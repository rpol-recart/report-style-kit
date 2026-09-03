# report-style-kit

## Назначение

Комплект для генерации отчётов в стиле образца: по отчёту прошлого квартала
(`.docx`) и материалам (`materials/`) формируются законченные `отчёт.docx`
и `отчёт.pdf` с гарантированно теми же шрифтами, оформлением, тоном и доменной
лексикой. Контур закрытый: модель пишет контент, скрипты гарантируют оформление —
стиль воспроизводится template-наследованием, блоковой грамматикой манифеста
и двухстадийным линтом.

## Пререквизиты

- Python 3 ≥ 3.10.
- Пакет `python-docx`. Проверка: `python3 -c "import docx"`.
- LibreOffice — опционально, только для PDF-рендера. Проверка: `command -v soffice`
  (без него конвейер даёт только .docx).

## Быстрый старт

Полный конвейер 0–7 (команды выполняются в каталоге комплекта; подробно —
`skills/report-kit-orchestrator/SKILL.md`):

```bash
# 0. Проверки
python3 -c "import docx"

# 1. Извлечь стиль из образца
python3 scripts/extract_profile.py --sample образец.docx --outdir work/

# 2. Выкурить styleguide.draft.md и glossary.draft.md → work/styleguide.md, work/glossary.md

# 3. Развернуть materials/ в манифест (по skills/report-content-plan/SKILL.md)

# 4. Линт манифеста (FAIL → править манифест, максимум 3 итерации)
python3 scripts/lint_text.py --manifest content-manifest.json --profile work/style-profile.json

# 5. Сборка docx
python3 scripts/build_report.py --manifest content-manifest.json --profile work/style-profile.json --template work/template.docx -o отчёт.docx

# 6. Линт собранного отчёта (FAIL → цикл правок манифест → сборка → линт, максимум 3)
python3 scripts/lint_text.py --docx отчёт.docx --blacklist work/slang-blacklist.txt
python3 scripts/lint_visual.py --docx отчёт.docx --profile work/style-profile.json

# 7. Рендер PDF
python3 scripts/render_pdf.py --docx отчёт.docx -o out/
```

## Структура каталога

```text
report-style-kit/
├── README.md                 # этот файл: входная точка комплекта
├── skills/                   # 5 скиллов-инструкций для модели
│   ├── report-kit-orchestrator/  # конвейер 0–7 целиком, STOP-условия, запреты
│   ├── report-style-extract/     # извлечение стиля из образца + курирование
│   ├── report-content-plan/      # разворот черновиков в content-manifest.json
│   ├── report-docx-build/        # сборка .docx из манифеста, коды выхода
│   └── report-style-lint/        # двухстадийный линт, категории FAIL/warn
├── scripts/                  # 5 детерминированных скриптов (python-docx)
│   ├── extract_profile.py    # образец → work/ (профиль, шаблон, черновики)
│   ├── build_report.py       # манифест + профиль + шаблон → отчёт.docx
│   ├── lint_text.py          # анти-сленг, placeholders, типографика
│   ├── lint_visual.py        # стили/шрифты/таблицы/поля против профиля
│   └── render_pdf.py         # .docx → .pdf через soffice headless
├── schemas/                  # контракты форматов
│   ├── content-manifest.schema.json  # блоковая грамматика (7 типов блоков)
│   └── style-profile.schema.json     # профиль стиля образца
├── base/
│   └── slang-blacklist.txt   # база анти-сленга: категория<TAB>regex
└── tests/                    # golden-тест, линт-фикстуры, синтетический прогон
```

## Развёртывание в контуре

Скопируйте каталог report-style-kit/ целиком в рабочую среду. Положите образец
отчёта (прошлый квартал) и materials/ рядом. Откройте
skills/report-kit-orchestrator/SKILL.md и следуйте шагам.

## Запуск агентов: последовательность и промпты

Готовые промпты для агента в контуре. Два режима: **А** — один прогон (агент проходит весь конвейер сам), **Б** — поэтапный запуск (контроль оператора между стадиями). Копируйте промпты дословно; рабочие каталоги — как в «Быстром старте».

### Режим А: один прогон (мастер-промпт)

```text
Ты — генератор квартального отчёта. Работай строго по инструкции
skills/report-kit-orchestrator/SKILL.md, не пропуская шаги.

Вход: ../sample.docx — образец отчёта прошлого квартала;
../materials/ — черновики, числа, графики с описаниями.

Выполни конвейер 0–7: извлечение стиля → курирование styleguide/glossary →
content-manifest.json (тексты блоков пиши по правилам
skills/report-technical-writer/SKILL.md, структуру — по
skills/report-content-plan/SKILL.md) → линт манифеста → сборка →
финальные линты → PDF.

Жёсткие правила: числа только из materials (нет числа — $unresolved);
после 3 неудачных итераций линта — стоп и вопрос; docx руками не править;
скрипты и template.docx не изменять; выдумывать запрещено.

Верни: пути report.docx и report.pdf; итог всех линтов (PASS/FAIL);
список незакрытых вопросов (если есть $unresolved).
```

### Режим Б: поэтапный запуск

| # | Стадия | Промпт агенту | Критерий перехода к следующей |
|---|--------|---------------|-------------------------------|
| 1 | Извлечение стиля | см. «Промпт Б1» | 5 артефактов в work/, exit 0 |
| 2 | Курирование стандарта | см. «Промпт Б2» | оператор подтвердил styleguide.md и glossary.md |
| 3 | Контент-план | см. «Промпт Б3» | content-manifest.json создан |
| 4 | Линт-контур | см. «Промпт Б4» | lint_text --manifest → PASS |
| 5 | Сборка и верификация | см. «Промпт Б5» | оба линта PASS, report.docx (+ report.pdf) exist |

**Промпт Б1 (извлечение):**

```text
Выполни: python3 scripts/extract_profile.py --sample ../sample.docx --outdir work/
Затем прочитай skills/report-style-extract/SKILL.md и проверь артефакты по
правилам курирования. Отчитайся: (1) title_block — какие кегли/жирность у титула;
(2) список имён table_templates с заголовками колонок; (3) первые 10 терминов
глоссария; (4) разделы section_blueprint. Ничего не курируй без моего указания.
```

**Промпт Б2 (курирование):**

```text
Прочитай work/styleguide.draft.md и work/glossary.draft.md. Курай их по правилам
skills/report-style-extract/SKILL.md: вычини мусорные токены глоссария, проверь
title_block по образцу, переименуй table_template_N в styleguide в осмысленные
имена (имена в style-profile.json не менять). Сохрани как work/styleguide.md и
work/glossary.md. Покажи diff-сводку изменений.
```

**Промпт Б3 (контент-план):**

```text
Прочитай ../materials/ (outline, черновики, числа, описания графиков),
work/styleguide.md, work/glossary.md. Построй work/content-manifest.json по
skills/report-content-plan/SKILL.md: структура разделов по section_blueprint,
тексты блоков по skills/report-technical-writer/SKILL.md. Числа ТОЛЬКО из
materials; отсутствующее помечай "$unresolved". Таблицы — по шаблонам из
style-profile.json (table_templates). Верни: число секций/блоков, список
использованных шаблонов таблиц, список $unresolved (если есть).
```

**Промпт Б4 (линт-контур):**

```text
Прогоняй: python3 scripts/lint_text.py --manifest work/content-manifest.json
--profile work/style-profile.json. При FAIL исправляй manifest (правила —
skills/report-technical-writer/SKILL.md) и повторяй. Максимум 3 итерации.
Отчитайся: финальный статус, что правил, оставшиеся предупреждения.
Если $unresolved остались — стоп, перечисли их вопросами ко мне.
```

**Промпт Б5 (сборка и верификация):**

```text
Выполни: python3 scripts/build_report.py --manifest work/content-manifest.json
--profile work/style-profile.json --template work/template.docx -o report.docx
Затем: lint_text.py --docx report.docx --blacklist work/slang-blacklist.txt и
lint_visual.py --docx report.docx --profile work/style-profile.json.
Оба должны быть PASS (до 3 итераций правок manifest → пересборка).
Затем: python3 scripts/render_pdf.py --docx report.docx -o .
Верни: пути и размеры report.docx/report.pdf, итог линтов.
```

### Правила для обоих режимов

- Промпты не изменять: в них зашиты лимиты итераций и STOP-условия
- Между стадиями режима Б оператор видит отчёт агента и подтверждает продолжение
- Любой STOP агента = вопрос оператору, а не самостоятельное решение

## Ограничения

- Закрытый контур: без интернета и внешних сервисов; всё нужное — в комплекте.
- Глоссарий строится только из образца: термины вне глоссария дают предупреждение
  линта (glossary-warn).
- Графики не генерируются: готовые PNG поставляются в `materials/charts/`
  вместе с описаниями.
- PDF требует установленного soffice; при его отсутствии выход — только .docx.
