"""
App web: lê a planilha Excel (prog.xlsm) e exibe preventivas da frota.
Colunas (Excel): A tipo, B frota, C OS, E data, G setor, I tipo plano, J status.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "prog.xlsm"

# Letras Excel (1-based) -> índice 0-based no DataFrame sem colunas extras à esquerda
COL_A = 0  # tipo equipamento
COL_B = 1  # código frota
COL_C = 2  # ordem de serviço
COL_E = 4  # data preventiva
COL_G = 6  # setor
COL_I = 8  # tipo do plano
COL_J = 9  # status (P/A/E)

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


def load_rows() -> tuple[list[dict], str | None]:
    """Carrega linhas da planilha; retorna (lista de dicts, mensagem de erro ou None)."""
    if not EXCEL_PATH.is_file():
        return [], f"Arquivo não encontrado: {EXCEL_PATH.name}"

    try:
        xl = pd.ExcelFile(EXCEL_PATH, engine="openpyxl")
        sheet = "WHATSAPP" if "WHATSAPP" in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet, header=None, engine="openpyxl")
    except Exception as e:
        return [], str(e)

    if df.shape[1] <= COL_J:
        return [], "Planilha sem colunas suficientes (até J)."

    # Cabeçalho costuma estar na linha 1 (índice 1); dados a partir da 2
    start = 2
    if df.shape[0] <= start:
        return [], "Nenhuma linha de dados."

    rows: list[dict] = []
    for i in range(start, len(df)):
        r = df.iloc[i]
        tipo_plano = r.iloc[COL_I]
        if pd.isna(tipo_plano) or str(tipo_plano).strip() == "":
            tipo_plano = r.iloc[3] if df.shape[1] > 3 else None  # coluna D PLANO

        data_prev = r.iloc[COL_E]
        if hasattr(data_prev, "strftime"):
            data_str = data_prev.strftime("%Y-%m-%d")
        elif pd.isna(data_prev):
            data_str = ""
        else:
            data_str = str(data_prev)[:10]

        st = _normalize_status(r.iloc[COL_J])
        rows.append(
            {
                "tipo_equipamento": _cell_str(r.iloc[COL_A]),
                "cod_frota": _cell_str(r.iloc[COL_B]),
                "ordem_servico": _cell_str(r.iloc[COL_C]),
                "data_preventiva": data_str,
                "setor": _cell_str(r.iloc[COL_G]),
                "tipo_plano": _cell_str(tipo_plano),
                "status": st,
                "status_label": STATUS_LABEL.get(st, st or "—"),
            }
        )

    return rows, None


_cache: dict | None = None  # {"mtime": float, "rows": list, "err": str | None}


def get_rows() -> tuple[list[dict], str | None, str]:
    """
    Fluxo de dados: recarrega do disco quando prog.xlsm muda (mtime).
    Mesma revisão = mesmos dados em memória (cache).
    """
    global _cache
    if not EXCEL_PATH.is_file():
        _cache = None
        return [], f"Arquivo não encontrado: {EXCEL_PATH.name}", "missing"

    try:
        mtime = EXCEL_PATH.stat().st_mtime
    except OSError:
        _cache = None
        return [], "Não foi possível acessar a planilha.", "error"

    revision = f"{mtime:.6f}"

    if _cache is not None and _cache["mtime"] == mtime:
        return _cache["rows"], _cache["err"], revision

    rows, err = load_rows()
    _cache = {"mtime": mtime, "rows": rows, "err": err}
    return rows, err, revision


def _cell_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val).strip()


app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False


@app.route("/")
def index():
    rows, err, revision = get_rows()
    return render_template(
        "index.html",
        rows=rows,
        error=err,
        excel_name=EXCEL_PATH.name,
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
    # Planilha é relida quando o arquivo muda (mtime); o front chama /api/dados em intervalos.
    app.run(host="127.0.0.1", port=5000, debug=True)
