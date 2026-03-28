"""
App web: lê a planilha Excel (prog.xlsm) e exibe preventivas da frota.
"""

from __future__ import annotations
from pathlib import Path
import pandas as pd
from flask import Flask, jsonify, render_template

# 🔥 CAMINHO CORRETO (raiz do projeto)
BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_PATH = BASE_DIR / "prog.xlsm"

# COLUNAS
COL_A = 0
COL_B = 1
COL_C = 2
COL_E = 4
COL_G = 6
COL_I = 8
COL_J = 9

STATUS_LABEL = {
    "P": "Programada",
    "A": "Andamento",
    "E": "Encerrada",
}


def _normalize_status(raw) -> str:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = str(raw).strip().upper()
    return s[0] if s else ""


def _cell_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val).strip()


def load_rows():
    if not EXCEL_PATH.is_file():
        return [], f"Arquivo não encontrado: {EXCEL_PATH}"

    try:
        df = pd.read_excel(EXCEL_PATH, engine="openpyxl", header=None)
    except Exception as e:
        return [], str(e)

    rows = []

    for i in range(2, len(df)):
        r = df.iloc[i]

        data_prev = r.iloc[COL_E]
        if hasattr(data_prev, "strftime"):
            data_str = data_prev.strftime("%Y-%m-%d")
        else:
            data_str = str(data_prev)[:10] if not pd.isna(data_prev) else ""

        st = _normalize_status(r.iloc[COL_J])

        rows.append(
            {
                "tipo_equipamento": _cell_str(r.iloc[COL_A]),
                "cod_frota": _cell_str(r.iloc[COL_B]),
                "ordem_servico": _cell_str(r.iloc[COL_C]),
                "data_preventiva": data_str,
                "setor": _cell_str(r.iloc[COL_G]),
                "tipo_plano": _cell_str(r.iloc[COL_I]),
                "status": st,
                "status_label": STATUS_LABEL.get(st, st),
            }
        )

    return rows, None


# 🔥 IMPORTANTE: apontando pro templates correto
app = Flask(__name__, template_folder="../templates")
app.config["JSON_AS_ASCII"] = False


@app.route("/")
def index():
    rows, err = load_rows()
    return render_template(
        "index.html",
        rows=rows,
        error=err,
        excel_name=EXCEL_PATH.name,
    )


@app.route("/api/dados")
def api_dados():
    rows, err = load_rows()
    return jsonify(
        {
            "ok": err is None,
            "error": err,
            "rows": rows if err is None else [],
        }
    )


# ⚠️ não é usado no Vercel, mas ok deixar
if __name__ == "__main__":
    app.run(debug=True)