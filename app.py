"""
App web: lê a planilha Excel e exibe ordens de serviço.

Usa o primeiro arquivo existente: prog.xlsm ou prog.xlsx (mesma pasta que app.py),
ou o caminho em PROG_EXCEL_PATH.

Colunas (Excel): A unidade, C nº boletim, D frota, E status, F data,
H tipo equipamento, K plano, L setor.
"""

from __future__ import annotations

import numbers
import os
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template

BASE_DIR = Path(__file__).resolve().parent
_override = (os.environ.get("PROG_EXCEL_PATH") or "").strip()


def resolve_excel_path() -> Path:
    """Caminho da planilha: variável de ambiente, ou prog.xlsm, ou prog.xlsx na pasta do projeto."""
    if _override:
        return Path(_override).expanduser().resolve()
    for name in ("prog.xlsm", "prog.xlsx"):
        p = (BASE_DIR / name).resolve()
        if p.is_file():
            return p
    return (BASE_DIR / "prog.xlsx").resolve()

# Índices 0-based (coluna A = 0)
COL_A = 0   # unidade
COL_C = 2   # número do boletim
COL_D = 3   # frota do equipamento
COL_E = 4   # status da ordem (P/A/E)
COL_F = 5   # data da ordem
COL_H = 7   # tipo do equipamento
COL_K = 10  # plano da ordem
COL_L = 11  # setor

COL_MAX = COL_L

STATUS_LABEL = {
    "P": "Programada",
    "A": "Andamento",
    "E": "Encerrada",
}


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
        raw = raw.strip()
        if not raw:
            return ""

    ts = pd.to_datetime(raw, dayfirst=True, errors="coerce")
    if pd.isna(ts) and isinstance(raw, numbers.Real) and not isinstance(raw, bool):
        ts = pd.to_datetime(float(raw), unit="d", origin="1899-12-30", errors="coerce")
    if pd.isna(ts):
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

    # Linha 1 do Excel = título das colunas (ignorada). Linha 2 = primeira linha de dados (pandas índice 1).
    start = 1
    if df.shape[0] <= start:
        return [], "Nenhuma linha de dados."

    rows: list[dict] = []
    for i in range(start, len(df)):
        r = df.iloc[i]
        tipo_plano = r.iloc[COL_K]
        data_str = _format_data_br(r.iloc[COL_F])

        st = _normalize_status(r.iloc[COL_E])
        rows.append(
            {
                "unidade": _cell_str(r.iloc[COL_A]),
                "numero_boletim": _cell_str(r.iloc[COL_C]),
                "cod_frota": _cell_str(r.iloc[COL_D]),
                "data_ordem": data_str,
                "tipo_equipamento": _cell_str(r.iloc[COL_H]),
                "tipo_plano": _cell_str(tipo_plano),
                "setor": _cell_str(r.iloc[COL_L]),
                "status": st,
                "status_label": STATUS_LABEL.get(st, st or "—"),
            }
        )

    return rows, None


_cache: dict | None = None  # path (str), mtime, rows, err


def get_rows() -> tuple[list[dict], str | None, str]:
    global _cache
    path = resolve_excel_path()
    if not path.is_file():
        _cache = None
        return (
            [],
            f"Nenhuma planilha encontrada. Coloque prog.xlsm ou prog.xlsx em: {BASE_DIR}",
            "missing",
        )

    try:
        mtime = path.stat().st_mtime
    except OSError:
        _cache = None
        return [], "Não foi possível acessar a planilha.", "error"

    pkey = str(path.resolve())
    revision = f"{mtime:.7f}:{path.name}"

    if (
        _cache is not None
        and _cache.get("path") == pkey
        and _cache.get("mtime") == mtime
    ):
        return _cache["rows"], _cache["err"], revision

    rows, err = load_rows(path)
    _cache = {"path": pkey, "mtime": mtime, "rows": rows, "err": err}
    return rows, err, revision


def _cell_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val).strip()


app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["JSON_AS_ASCII"] = False


@app.route("/")
def index():
    path = resolve_excel_path()
    rows, err, revision = get_rows()
    excel_name = path.name if path.is_file() else "prog.xlsm ou prog.xlsx"
    return render_template(
        "index.html",
        rows=rows,
        error=err,
        excel_name=excel_name,
        revision=revision,
    )


@app.route("/api/dados")
def api_dados():
    rows, err, revision = get_rows()
    out = jsonify(
        {
            "ok": err is None,
            "error": err,
            "rows": rows if err is None else [],
            "revision": revision,
        }
    )
    out.headers["Cache-Control"] = "no-store"
    return out


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
