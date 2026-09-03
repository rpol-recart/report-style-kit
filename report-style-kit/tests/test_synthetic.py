#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_synthetic.py — Task 9.3 комплекта report-style-kit. Синтетический прогон
materials_synthetic/ → docx + pdf полным конвейером (subprocess'ы, cwd — корень
комплекта, аргументы относительные).

Шаги:
  1. extract_profile.py по эталону → tests/work_synthetic/ (профиль + шаблон);
  2. lint_text.py --manifest tests/manifest_synthetic.json → при FAIL сборка
     ЗАПРЕЩЕНА (exit 1);
  3. build_report.py → tests/out_synthetic/report.docx;
  4. lint_text.py --docx report.docx → обязан быть зелёным;
  5. lint_visual.py report.docx против профиля → обязан быть зелёным;
  6. render_pdf.py → tests/out_synthetic/report.pdf.

SYNTHETIC PASS (exit 0), если все шаги зелёные И report.docx существует и
> 10 КБ, И report.pdf существует и > 5 КБ. Любой сбой — отчёт и exit 1.

Запуск: python3 tests/test_synthetic.py
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

KIT = Path(__file__).resolve().parents[1]
SAMPLE_REL = "../CR_Predictor_Report.docx"
WORK_REL = "tests/work_synthetic"
MANIFEST_REL = "tests/manifest_synthetic.json"
OUT_REL = "tests/out_synthetic"
DOCX_REL = f"{OUT_REL}/report.docx"
PDF_REL = f"{OUT_REL}/report.pdf"
MIN_DOCX = 10 * 1024   # байт
MIN_PDF = 5 * 1024     # байт


def step(cmd: list, what: str, expect_rc: int = 0) -> subprocess.CompletedProcess:
    proc = subprocess.run([sys.executable] + cmd, cwd=KIT,
                          capture_output=True, text=True)
    if proc.returncode != expect_rc:
        print(f"[synthetic] шаг провален: {what} "
              f"(exit {proc.returncode}, ожидался {expect_rc})", file=sys.stderr)
        print("--- stdout ---\n" + proc.stdout, file=sys.stderr)
        print("--- stderr ---\n" + proc.stderr, file=sys.stderr)
        raise SystemExit(1)
    return proc


def main() -> int:
    (KIT / WORK_REL).mkdir(parents=True, exist_ok=True)
    (KIT / OUT_REL).mkdir(parents=True, exist_ok=True)

    # 1. профиль + шаблон из эталона
    step(["scripts/extract_profile.py", "--sample", SAMPLE_REL,
          "--outdir", WORK_REL], "extract_profile.py")

    # 2. lint текста манифеста — BUILD-ЗАПРЕТ при FAIL
    step(["scripts/lint_text.py", "--manifest", MANIFEST_REL,
          "--profile", f"{WORK_REL}/style-profile.json"],
         "lint_text --manifest")

    # 3. сборка
    step(["scripts/build_report.py", "--manifest", MANIFEST_REL,
          "--profile", f"{WORK_REL}/style-profile.json",
          "--template", f"{WORK_REL}/template.docx",
          "-o", DOCX_REL], "build_report.py")

    # 4. lint текста собранного docx
    step(["scripts/lint_text.py", "--docx", DOCX_REL,
          "--profile", f"{WORK_REL}/style-profile.json"],
         "lint_text --docx")

    # 5. визуальный линт
    proc_vis = step(["scripts/lint_visual.py", "--docx", DOCX_REL,
                     "--profile", f"{WORK_REL}/style-profile.json"],
                    "lint_visual")

    # 6. PDF
    proc_pdf = step(["scripts/render_pdf.py", "--docx", DOCX_REL,
                     "-o", OUT_REL], "render_pdf")

    # 7. двойной выход: существование и размеры
    problems = []
    docx_path = KIT / DOCX_REL
    pdf_path = KIT / PDF_REL
    if not docx_path.is_file():
        problems.append(f"{DOCX_REL} не существует")
    elif docx_path.stat().st_size <= MIN_DOCX:
        problems.append(f"{DOCX_REL} размер {docx_path.stat().st_size} байт "
                        f"≤ {MIN_DOCX}")
    if not pdf_path.is_file():
        problems.append(f"{PDF_REL} не существует")
    elif pdf_path.stat().st_size <= MIN_PDF:
        problems.append(f"{PDF_REL} размер {pdf_path.stat().st_size} байт "
                        f"≤ {MIN_PDF}")
    if problems:
        print("[synthetic] проблемы с выходными файлами:", file=sys.stderr)
        for p in problems:
            print("  -", p, file=sys.stderr)
        return 1

    print("SYNTHETIC PASS")
    print("lint_visual:", proc_vis.stdout.strip().splitlines()[-1])
    print("render_pdf:", proc_pdf.stdout.strip().splitlines()[-1])
    print(f"report.docx: {docx_path.stat().st_size} байт, "
          f"report.pdf: {pdf_path.stat().st_size} байт")
    return 0


if __name__ == "__main__":
    sys.exit(main())
