from __future__ import annotations

import argparse
import hashlib
import os
import tempfile
import urllib.error
import urllib.request
from pathlib import Path


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def download_to_temp(url: str, timeout: int) -> Path:
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (compatible; excel-sync-bot/1.0)"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = resp.read()
    if not data:
        raise RuntimeError("Download da planilha retornou vazio.")

    fd, tmp_name = tempfile.mkstemp(prefix="excel-sync-", suffix=".xlsx")
    os.close(fd)
    tmp = Path(tmp_name)
    tmp.write_bytes(data)
    return tmp


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Baixa planilha da URL e atualiza arquivo local se houver alteração."
    )
    parser.add_argument("--url", required=True, help="URL direta do arquivo xlsx")
    parser.add_argument(
        "--output",
        default="prog.xlsx",
        help="Caminho de saída da planilha no repositório (padrão: prog.xlsx)",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=90,
        help="Timeout do download em segundos (padrão: 90)",
    )
    args = parser.parse_args()

    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    try:
        tmp = download_to_temp(args.url, timeout=args.timeout)
    except (urllib.error.URLError, TimeoutError, ValueError) as e:
        raise RuntimeError(f"Falha ao baixar planilha: {e}") from e

    try:
        new_hash = sha256_file(tmp)
        old_hash = sha256_file(output) if output.is_file() else None
        if old_hash == new_hash:
            print("Sem alterações no Excel.")
            return 0
        tmp.replace(output)
        print(f"Planilha atualizada: {output}")
        return 0
    finally:
        if tmp.exists():
            tmp.unlink(missing_ok=True)


if __name__ == "__main__":
    raise SystemExit(main())
