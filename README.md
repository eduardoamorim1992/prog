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

### Limitação importante (planilha)

No servidor serverless **não há** o arquivo Excel da sua máquina. Sem o ficheiro, a app mostra “arquivo não encontrado”. Opções:

- **Teste rápido:** inclua um `prog.xlsx` pequeno no repositório (remova `*.xlsx` do `.gitignore` só se puder versionar dados).
- **Produção:** hospedar o ficheiro acessível por URL e estender o código para ler dessa URL, ou usar **Render / Railway / Fly.io** com disco persistente.

Variável opcional em produção: `PROG_EXCEL_PATH` (caminho absoluto **dentro** do bundle deployado, ex.: `/var/task/prog.xlsx`, se copiar o ficheiro no build).

## Variáveis de ambiente

| Variável | Uso |
|----------|-----|
| `PROG_EXCEL_PATH` | Caminho absoluto da planilha (local ou deploy). |
