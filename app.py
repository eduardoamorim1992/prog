"""
App web: lê a planilha Excel e exibe ordens de serviço.

Fontes da planilha (por ordem):
- PROG_EXCEL_PATH — caminho absoluto (prioridade se definido)
- prog.xlsm / prog.xlsx na pasta do projeto (incl. deploy Vercel no bundle)
- PROG_EXCEL_URL — download público se não houver ficheiro local

Atualização online (URL): defina PROG_EXCEL_REFRESH_SECS (ex.: 3600 = nova tentativa de download a cada 1 hora).

Colunas (planilha principal): A unidade, B boletim, C frota, D status, E data,
F descrição (códigos do plano de manutenção), H tipo/grupo equipamento,
J chave (liga à aba base), K plano, L setor.

Aba **base**: coluna D = descrição do serviço, coluna E = chave (igual à J da principal).
"""

from __future__ import annotations

import numbers
import os
import re
import time
import urllib.error
import urllib.request
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template

BASE_DIR = Path(__file__).resolve().parent
_override = (os.environ.get("PROG_EXCEL_PATH") or "").strip()
_excel_url = (os.environ.get("PROG_EXCEL_URL") or "").strip()


def _env_int(name: str, default: int, *, min_v: int | None = None, max_v: int | None = None) -> int:
    try:
        v = int(os.environ.get(name, str(default)))
    except ValueError:
        v = default
    if min_v is not None:
        v = max(min_v, v)
    if max_v is not None:
        v = min(max_v, v)
    return v

# Cache em /tmp no Vercel (serverless)
_URL_TMP = Path("/tmp/prog_frota_planilha.xlsx")
_url_last_fetch = 0.0


def _ensure_url_excel(url: str) -> tuple[Path | None, str | None]:
    """Garante ficheiro local em /tmp baixado da URL. Retorna (path, erro)."""
    global _url_last_fetch

    refresh = _env_int("PROG_EXCEL_REFRESH_SECS", 90, min_v=15, max_v=86_400)
    now = time.time()

    if _URL_TMP.is_file() and (now - _url_last_fetch) < refresh:
        return _URL_TMP, None

    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; prog-frota/1.0)"},
        )
        with urllib.request.urlopen(req, timeout=90) as resp:
            data = resp.read()
        if not data:
            if _URL_TMP.is_file():
                return _URL_TMP, None
            return None, "A URL da planilha retornou conteúdo vazio."
        _URL_TMP.write_bytes(data)
        _url_last_fetch = now
    except (urllib.error.URLError, OSError, TimeoutError, ValueError) as e:
        if _URL_TMP.is_file():
            return _URL_TMP, None
        return None, f"Erro ao baixar a planilha: {e}"

    return _URL_TMP, None


def resolve_excel_path() -> Path:
    """Caminho esperado (para mensagens); pode não existir se nada configurado."""
    if _override:
        return Path(_override).expanduser().resolve()
    for name in ("prog.xlsm", "prog.xlsx"):
        p = (BASE_DIR / name).resolve()
        if p.is_file():
            return p
    if _excel_url:
        return _URL_TMP
    return (BASE_DIR / "prog.xlsx").resolve()


def _resolve_readable_excel() -> tuple[Path | None, str | None]:
    """Path de um ficheiro que existe e pode ser lido, ou (None, erro)."""
    if _override:
        p = Path(_override).expanduser().resolve()
        return p, None
    for name in ("prog.xlsm", "prog.xlsx"):
        p = (BASE_DIR / name).resolve()
        if p.is_file():
            return p, None
    if _excel_url:
        return _ensure_url_excel(_excel_url)
    return (BASE_DIR / "prog.xlsx").resolve(), None


def _is_url_cache_path(path: Path) -> bool:
    try:
        return path.resolve() == _URL_TMP.resolve()
    except OSError:
        return False


# Índices 0-based (coluna A = 0). Planilha atual: B BOLETIM, C FROTA, D STATUS, E DATA.
COL_A = 0   # A unidade (INSTÂNCIA)
COL_B = 1   # B número do boletim
COL_C = 2   # C frota
COL_D = 3   # D status (P/A/E)
COL_E = 4   # E data da ordem
COL_F = 5   # F descrição (texto com códigos após "Manutenção:")
COL_H = 7   # H grupo / tipo equipamento
COL_J = 9   # J chave → aba base col. E
COL_K = 10  # K plano
COL_L = 11  # L setor

COL_MAX = COL_L

# Nome amigável por código da instância (coluna A). Códigos não listados seguem o texto da planilha.
UNIDADE_NOMES: dict[str, str] = {
    "USA1": "Unidade Iguatemi",
    "USA2": "Unidade Paranacity",
    "USA3": "Unidade Tapejara",
    "USA4": "Unidade Ivate",
    "USA13": "Unidade Terra Rica",
    "USA15": "Unidade Rondon",
    "USA16": "Unidade Cidade Gaucha",
    "USA17": "Unidade URP",
    "USA18": "Unidade Moreira Sales",
}


def _unidade_exibicao(codigo: str) -> str:
    c = (codigo or "").strip().upper()
    if not c:
        return ""
    return UNIDADE_NOMES.get(c, (codigo or "").strip())


# Aba base: serviços por chave
BASE_COL_D = 3  # texto do serviço
BASE_COL_E = 4  # chave (mesmo valor que coluna J da principal)

STATUS_LABEL = {
    "P": "Programada",
    "A": "Andamento",
    "E": "Encerrada",
}


def _extrair_codigos_plano_descricao(val) -> str:
    """
    Extrai trecho numérico após 'Manutenção:' em textos do tipo
    '... Plano(s) de Manutenção: 110040000/110020000/100010000**(...)...'
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if not s:
        return ""
    if "_x" in s.lower():
        s = _clean_excel_escapes(s)
    m = re.search(
        r"Manuten[çc][aã]o\s*:\s*([0-9]+(?:\s*/\s*[0-9]+)*)",
        s,
        re.IGNORECASE,
    )
    if not m:
        m = re.search(
            r"Manutencao\s*:\s*([0-9]+(?:\s*/\s*[0-9]+)*)",
            s,
            re.IGNORECASE,
        )
    if not m:
        return ""
    return re.sub(r"\s+", "", m.group(1))


def _clean_excel_escapes(s: str) -> str:
    """Remove códigos tipo _x000D_ (CR/LF do Excel) e normaliza espaços."""
    if not s:
        return s
    t = re.sub(r"_x000D_", " ", s, flags=re.I)
    t = re.sub(r"_x000A_", " ", t, flags=re.I)
    t = re.sub(r"_x0009_", " ", t, flags=re.I)
    return re.sub(r"\s+", " ", t).strip()


def _normalize_status(raw) -> str:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = str(raw).strip().upper()
    if not s:
        return ""
    return s[0] if s[0] in STATUS_LABEL else s


def _format_data_br(raw) -> str:
    """Data da ordem exibida como DD.MM.AAAA (dia.mês.ano)."""
    if raw is None:
        return ""
    if isinstance(raw, float) and pd.isna(raw):
        return ""

    if isinstance(raw, str):
        raw = _clean_excel_escapes(raw.strip())
        if not raw:
            return ""

    ts = pd.to_datetime(raw, dayfirst=True, errors="coerce")
    if pd.isna(ts) and isinstance(raw, numbers.Real) and not isinstance(raw, bool):
        ts = pd.to_datetime(float(raw), unit="d", origin="1899-12-30", errors="coerce")
    if pd.isna(ts):
        if isinstance(raw, str):
            return raw
        return str(raw).strip()

    return ts.strftime("%d.%m.%Y")


def load_rows(path: Path) -> tuple[list[dict], str | None]:
    """Carrega linhas da planilha; retorna (lista de dicts, mensagem de erro ou None)."""
    if not path.is_file():
        return [], f"Arquivo não encontrado: {path}"

    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
        sheet = "WHATSAPP" if "WHATSAPP" in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
    except Exception as e:
        return [], str(e)

    if df.shape[1] <= COL_MAX:
        return [], "Planilha sem colunas suficientes (necessário até a coluna L)."

    start = 1
    if df.shape[0] <= start:
        return [], "Nenhuma linha de dados."

    rows: list[dict] = []
    for i in range(start, len(df)):
        r = df.iloc[i]
        tipo_plano = r.iloc[COL_K]
        data_str = _format_data_br(r.iloc[COL_E])
        plano_codigos = _extrair_codigos_plano_descricao(r.iloc[COL_F])

        st = _normalize_status(r.iloc[COL_D])
        cod_unidade = _cell_str(r.iloc[COL_A])
        rows.append(
            {
                "unidade": _unidade_exibicao(cod_unidade),
                "unidade_codigo": cod_unidade,
                "numero_boletim": _cell_str(r.iloc[COL_B]),
                "cod_frota": _cell_str(r.iloc[COL_C]),
                "data_ordem": data_str,
                "tipo_equipamento": _cell_str(r.iloc[COL_H]),
                "tipo_plano": _cell_str(tipo_plano),
                "plano_codigos": plano_codigos,
                "setor": _cell_str(r.iloc[COL_L]),
                "chave_os": _cell_str(r.iloc[COL_J]),
                "status": st,
                "status_label": STATUS_LABEL.get(st, st or "—"),
            }
        )

    return rows, None


def load_servicos_por_chave(path: Path) -> dict[str, list[str]]:
    """Lê aba 'base': coluna D serviço, E chave. Várias linhas por chave."""
    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
    except Exception:
        return {}

    sheet_name: str | None = None
    for name in xl.sheet_names:
        if name.strip().lower() == "base":
            sheet_name = name
            break
    if not sheet_name:
        return {}

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")
    except Exception:
        return {}

    if df.shape[1] <= BASE_COL_E:
        return {}

    start = 1
    if df.shape[0] <= start:
        return {}

    out: dict[str, list[str]] = {}
    for i in range(start, len(df)):
        row = df.iloc[i]
        chave = _cell_str(row.iloc[BASE_COL_E])
        if not chave:
            continue
        serv = _cell_str(row.iloc[BASE_COL_D])
        if chave not in out:
            out[chave] = []
        if serv:
            out[chave].append(serv)

    return out


_cache: dict | None = None


def get_rows() -> tuple[list[dict], str | None, str, dict[str, list[str]]]:
    global _cache
    path, prep_err = _resolve_readable_excel()
    if prep_err:
        _cache = None
        return [], prep_err, "missing", {}

    if path is None or not path.is_file():
        _cache = None
        extra = ""
        if os.environ.get("VERCEL"):
            extra = (
                " Inclua prog.xlsx no repositório ou defina PROG_EXCEL_URL (link direto ao .xlsx)."
            )
        return (
            [],
            f"Nenhuma planilha encontrada em {BASE_DIR}.{extra}",
            "missing",
            {},
        )

    try:
        mtime = path.stat().st_mtime
    except OSError:
        _cache = None
        return [], "Não foi possível acessar a planilha.", "error", {}

    pkey = str(path.resolve())
    from_url = _is_url_cache_path(path)
    if from_url:
        revision = f"url:{mtime:.4f}:{_url_last_fetch:.0f}"
    else:
        revision = f"{mtime:.7f}:{path.name}"

    if _cache is not None and _cache.get("path") == pkey and _cache.get("mtime") == mtime:
        if from_url:
            if _cache.get("url_fetch") == _url_last_fetch:
                return (
                    _cache["rows"],
                    _cache["err"],
                    revision,
                    _cache.get("servicos_por_chave", {}),
                )
        else:
            return (
                _cache["rows"],
                _cache["err"],
                revision,
                _cache.get("servicos_por_chave", {}),
            )

    rows, err = load_rows(path)
    servicos: dict[str, list[str]] = {}
    if err is None:
        servicos = load_servicos_por_chave(path)

    entry = {
        "path": pkey,
        "mtime": mtime,
        "rows": rows,
        "err": err,
        "servicos_por_chave": servicos,
    }
    if from_url:
        entry["url_fetch"] = _url_last_fetch
    _cache = entry
    return rows, err, revision, servicos


def _cell_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    s = str(val).strip()
    if "_x" in s.lower():
        s = _clean_excel_escapes(s)
    return s


app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["JSON_AS_ASCII"] = False


@app.route("/")
def index():
    path = resolve_excel_path()
    rows, err, revision, servicos_por_chave = get_rows()
    apath, _ = _resolve_readable_excel()
    if apath and apath.is_file() and _is_url_cache_path(apath):
        excel_name = "Planilha (URL)"
    else:
        excel_name = path.name if path.is_file() else "prog.xlsm ou prog.xlsx"
    return render_template(
        "index.html",
        rows=rows,
        error=err,
        excel_name=excel_name,
        revision=revision,
        servicos_por_chave=servicos_por_chave,
        data_poll_ms=_env_int("DATA_POLL_MS", 10_000, min_v=3_000, max_v=3_600_000),
        data_poll_hidden_ms=_env_int("DATA_POLL_HIDDEN_MS", 45_000, min_v=5_000, max_v=3_600_000),
    )


@app.route("/dashboard")
def dashboard():
    return render_template(
        "dashboard.html",
        dash_poll_ms=_env_int("DASH_POLL_MS", 3_600_000, min_v=0, max_v=86_400_000),
    )


@app.route("/api/dados")
def api_dados():
    rows, err, revision, servicos_por_chave = get_rows()
    payload = {
        "ok": err is None,
        "error": err,
        "rows": rows if err is None else [],
        "revision": revision,
        "servicos_por_chave": servicos_por_chave if err is None else {},
        "server_time": int(time.time()),
    }
    if _excel_url:
        payload["excel_refresh_secs"] = _env_int(
            "PROG_EXCEL_REFRESH_SECS", 90, min_v=15, max_v=86_400
        )
        payload["excel_source"] = "url"
        payload["last_excel_fetch_ts"] = int(_url_last_fetch) if _url_last_fetch else None
    else:
        payload["excel_source"] = "local"
    out = jsonify(payload)
    out.headers["Cache-Control"] = "no-store"
    return out


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
