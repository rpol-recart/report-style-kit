#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
render_pdf.py — Task 7 (Step 7.2) комплекта report-style-kit.

Конвертирует .docx в .pdf через LibreOffice headless (soffice). Даёт двойной
выход конвейера: report.docx + report.pdf.

CLI:
    python3 render_pdf.py --docx d.docx -o outdir/

Логика:
  1. Ищем исполняемый файл: shutil.which('soffice'), затем ('libreoffice').
     Оба отсутствуют → печать «PDF недоступен: soffice не найден. Отчёт
     сохранён только в .docx» и exit 3 (PDF — опциональный выход).
  2. Иначе subprocess: soffice --headless --convert-to pdf --outdir <dir>
     <docx>, таймаут 120 с.
  3. Проверяем, что <имя>.pdf существует и его размер > 0 → печать
     «PDF: <путь> (<N> байт)» и exit 0.
  4. Ошибка конвертации (ненулевой returncode, файл не создан, таймаут) →
     хвост stderr в stderr и exit 5. Код 2 — ошибки использования/IO.

Зависимости: только stdlib (subprocess, shutil). Python 3.12.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

TIMEOUT_S = 120
STDERR_TAIL = 500


def find_soffice() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Конвертация .docx в .pdf через LibreOffice headless.")
    ap.add_argument("--docx", required=True, help="исходный .docx")
    ap.add_argument("-o", "--outdir", required=True,
                    help="каталог для .pdf (создаётся при отсутствии)")
    args = ap.parse_args(argv)

    src = Path(args.docx)
    if not src.is_file():
        print(f"[ошибка] docx не найден: {src}", file=sys.stderr)
        return 2
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    soffice = find_soffice()
    if not soffice:
        print("PDF недоступен: soffice не найден. "
              "Отчёт сохранён только в .docx")
        return 3

    cmd = [soffice, "--headless", "--convert-to", "pdf",
           "--outdir", str(outdir), str(src)]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True,
                              timeout=TIMEOUT_S)
    except subprocess.TimeoutExpired:
        print(f"[ошибка] конвертация не уложилась в {TIMEOUT_S} с: "
              f"{' '.join(cmd)}", file=sys.stderr)
        return 5
    except OSError as exc:
        print(f"[ошибка] запуск {soffice} не удался: {exc}", file=sys.stderr)
        return 5

    pdf = outdir / (src.stem + ".pdf")
    if proc.returncode != 0 or not pdf.is_file() or pdf.stat().st_size == 0:
        tail = ((proc.stderr or "") + (proc.stdout or "")).strip()[-STDERR_TAIL:]
        print(f"[ошибка] конвертация в PDF не удалась "
              f"(returncode={proc.returncode}, pdf="
              f"{'создан' if pdf.is_file() else 'не создан'}).", file=sys.stderr)
        if tail:
            print("— хвост вывода конвертера —", file=sys.stderr)
            print(tail, file=sys.stderr)
        return 5

    print(f"PDF: {pdf} ({pdf.stat().st_size} байт)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
