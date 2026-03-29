# Ordens de serviço — preventivas (Flask)

Painel que lê `prog.xlsm` ou `prog.xlsx` (colunas A–L), filtros, gráficos de aderência e sincronização periódica no navegador.

## Local

```bash
pip install -r requirements.txt
python app.py
```

Abra `http://127.0.0.1:5000`. Coloque a planilha na mesma pasta que `app.py` ou defina `PROG_EXCEL_PATH` com o caminho absoluto do arquivo.

## Deploy no Vercel

O Vercel detecta o Flask em `app.py` ([documentação](https://vercel.com/docs/frameworks/backend/flask)).

1. Envie o repositório para o GitHub.
2. No [Vercel](https://vercel.com/new): **Add New Project** → importe o repo.
3. Framework: Flask (automático). **Deploy**.

### Planilha no Vercel (obrigatório: URL)

No Vercel não existe o teu `prog.xlsx` local. Define no projeto Vercel **Settings → Environment Variables**:

| Variável | Descrição |
|----------|-----------|
| **`PROG_EXCEL_URL`** | Link **direto** ao ficheiro `.xlsx` (o browser tem de conseguir descarregar sem login). Ex.: **GitHub** (*Raw* do ficheiro), **Google Drive** (link de download direto), **OneDrive** (partilha com link de download). |
| `PROG_EXCEL_REFRESH_SECS` | (Opcional) Segundos entre novos downloads. Predefinido: `90`. |

Depois de guardar as variáveis, faz **Redeploy** do projeto.

Outras opções: commitar um `prog.xlsx` no repo (tirar `*.xlsx` do `.gitignore` se for aceitável) ou usar `PROG_EXCEL_PATH` só quando o ficheiro existir no bundle.

## Variáveis de ambiente

| Variável | Uso |
|----------|-----|
| `PROG_EXCEL_URL` | URL HTTP(S) pública do `.xlsx` (recomendado no Vercel). |
| `PROG_EXCEL_PATH` | Caminho absoluto local (PC ou ficheiro copiado no deploy). |
| `PROG_EXCEL_REFRESH_SECS` | Intervalo de atualização ao usar URL. |
