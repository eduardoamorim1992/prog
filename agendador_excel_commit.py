import argparse
import datetime as dt
import subprocess
import sys
import time
from pathlib import Path

try:
    import win32com.client  # type: ignore[import-untyped]
except ImportError:
    win32com = None


def run_git_command(repo_path: Path, args: list[str]) -> tuple[int, str]:
    process = subprocess.run(
        ["git", *args],
        cwd=repo_path,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    output = (process.stdout or "") + (process.stderr or "")
    return process.returncode, output.strip()


def ensure_git_repository(repo_path: Path) -> None:
    code, output = run_git_command(repo_path, ["rev-parse", "--is-inside-work-tree"])
    if code != 0 or "true" not in output.lower():
        raise RuntimeError(f"Pasta nao e um repositorio Git: {repo_path}")


def refresh_excel_queries_with_excel_app(
    excel_path: Path, query_wait_seconds: int = 360
) -> None:
    if win32com is None:
        raise RuntimeError(
            "Dependencia ausente: pywin32. Instale com: pip install pywin32"
        )

    excel_app = win32com.client.DispatchEx("Excel.Application")
    workbook = None
    try:
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        workbook = excel_app.Workbooks.Open(str(excel_path))
        if workbook.ReadOnly:
            raise RuntimeError(
                "O arquivo foi aberto como somente leitura. "
                "Feche o Excel/OneDrive que esteja usando a planilha e tente novamente."
            )
        workbook.RefreshAll()
        print(f"Aguardando {query_wait_seconds}s para concluir atualizacao de queries...")
        time.sleep(query_wait_seconds)

        try:
            excel_app.CalculateUntilAsyncQueriesDone()
        except Exception:  # noqa: BLE001
            pass

        write_scheduler_stamp(workbook)
        workbook.Save()
        print("Excel atualizado e salvo com sucesso.")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel_app.Quit()


def write_scheduler_stamp(workbook) -> None:
    sheet_name = "__agendador_log"
    try:
        log_sheet = workbook.Worksheets(sheet_name)
    except Exception:  # noqa: BLE001
        log_sheet = workbook.Worksheets.Add()
        log_sheet.Name = sheet_name
        log_sheet.Cells(1, 1).Value = "ultima_execucao"
        log_sheet.Visible = 0  # xlSheetHidden

    log_sheet.Cells(2, 1).Value = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def commit_excel(repo_path: Path, excel_path: Path) -> None:
    rel_excel_path = excel_path.relative_to(repo_path)
    run_git_command(repo_path, ["add", str(rel_excel_path)])

    code, output = run_git_command(repo_path, ["diff", "--cached", "--quiet"])
    if code == 0:
        print("Sem alteracoes para commit.")
        return

    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    message = f"chore: atualizacao automatica do excel ({now})"

    code, output = run_git_command(repo_path, ["commit", "-m", message])
    if code != 0:
        raise RuntimeError(f"Falha ao criar commit.\n{output}")
    print(f"Commit criado com sucesso: {message}")

    code, output = run_git_command(repo_path, ["push", "origin", "main"])
    if code != 0:
        raise RuntimeError(f"Falha ao fazer push.\n{output}")
    print("Push realizado com sucesso.")


def run_scheduler(
    repo_path: Path,
    excel_relative_path: str,
    interval_seconds: int,
    query_wait_seconds: int,
) -> None:
    excel_path = repo_path / excel_relative_path
    ensure_git_repository(repo_path)
    if not excel_path.exists():
        raise RuntimeError(
            f"Arquivo Excel nao encontrado: {excel_path}. "
            "Informe o caminho correto com --excel-path."
        )

    print(f"Iniciando agendador. Repositorio: {repo_path}")
    print(f"Arquivo Excel: {excel_path}")
    print(f"Intervalo: {interval_seconds} segundos")

    while True:
        start = time.time()
        print(f"\n--- Ciclo iniciado em {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")
        try:
            refresh_excel_queries_with_excel_app(
                excel_path, query_wait_seconds=query_wait_seconds
            )
            commit_excel(repo_path, excel_path)
        except Exception as err:  # noqa: BLE001
            print(f"Erro no ciclo: {err}")

        elapsed = time.time() - start
        sleep_for = max(1, interval_seconds - int(elapsed))
        print(f"Aguardando {sleep_for} segundos para o proximo ciclo...")
        time.sleep(sleep_for)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Atualiza um Excel e faz commit automaticamente em intervalos fixos."
    )
    parser.add_argument(
        "--repo-path",
        default=".",
        help="Caminho do repositorio Git (padrao: pasta atual).",
    )
    parser.add_argument(
        "--excel-path",
        default="prog.xlsx",
        help="Caminho do arquivo Excel relativo ao repositorio.",
    )
    parser.add_argument(
        "--interval-seconds",
        type=int,
        default=3600,
        help="Intervalo entre ciclos em segundos (padrao: 3600 = 1 hora).",
    )
    parser.add_argument(
        "--query-wait-seconds",
        type=int,
        default=360,
        help="Tempo de espera para atualizacao das queries do Excel (padrao: 360).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    repo_path = Path(args.repo_path).resolve()
    run_scheduler(
        repo_path, args.excel_path, args.interval_seconds, args.query_wait_seconds
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())