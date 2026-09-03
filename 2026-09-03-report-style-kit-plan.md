# report-style-kit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use on-demands:subagent-driven-development (recommended) or on-demands:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Построить автономный комплект report-style-kit в /root/docs_template/report-style-kit/, воспроизводящий стиль эталонного отчёта в новых .docx/.pdf отчётах.

**Architecture:** 5 скиллов-инструкций (рус.) + 5 детерминированных скриптов (python-docx) + 2 JSON-схемы + база анти-сленга + тесты. Стиль гарантируется template-наследованием, блоковой грамматикой и двухстадийным линтом.

**Tech Stack:** Python 3.12, python-docx 1.2.0, JSON Schema (draft-07), LibreOffice headless (только PDF-рендер).

**Workflow prefs:** autoCommit=false, autoPush=false — шагов коммита нет; все записи локальные.

## File Structure

| Файл | Задача |
|---|---|
| report-style-kit/README.md | Task 1, Task 8 |
| report-style-kit/schemas/content-manifest.schema.json | Task 2 |
| report-style-kit/schemas/style-profile.schema.json | Task 2 |
| report-style-kit/base/slang-blacklist.txt | Task 3 |
| report-style-kit/scripts/extract_profile.py | Task 4 |
| report-style-kit/scripts/build_report.py | Task 5 |
| report-style-kit/scripts/lint_text.py | Task 6 |
| report-style-kit/scripts/lint_visual.py | Task 7 |
| report-style-kit/scripts/render_pdf.py | Task 7 |
| report-style-kit/skills/*/SKILL.md (5 шт.) | Task 8 |
| report-style-kit/tests/test_golden.py | Task 9 |
| report-style-kit/tests/test_lint_fixtures.py | Task 9 |
| report-style-kit/tests/materials_synthetic/ | Task 9 |

Не трогаем (scope fence): CR_Predictor_Report.docx (только чтение), всё вне /root/docs_template/, реестр on_demands.

---

## Task 1: Каркас комплекта

- [ ] **Step 1.1: Создать дерево каталогов**
  - **What:** mkdir -p /root/docs_template/report-style-kit/{skills/{report-kit-orchestrator,report-style-extract,report-content-plan,report-docx-build,report-style-lint},scripts,schemas,base,tests/materials_synthetic/charts}
  - **Why:** каркас для всех артефактов Goal.
  - **Inputs:** нет.
  - **Expected output:** пустое дерево каталогов.
  - **Verification:** `find /root/docs_template/report-style-kit -type d` показывает 11 каталогов.

- [ ] **Step 1.2: Заглушка README.md**
  - **What:** создать README.md с разделами: Назначение, Пререквизиты (python3 ≥3.10, python-docx, soffice опц.), Быстрый старт (конвейер 0–7), Структура каталога. Текст заполнится в Task 8.
  - **Why:** входная точка комплекта.
  - **Inputs:** дерево из 1.1.
  - **Expected output:** README.md с 4 разделами-заголовками.
  - **Verification:** `grep -c '^##' README.md` ≥ 4.

## Task 2: JSON-схемы

- [ ] **Step 2.1: content-manifest.schema.json**
  - **What:** схема: {meta:{title,subtitle,project,version,period}, sections:[{heading, level∈[1,2,3], blocks:[…]}]}. Типы блоков: thesis{text}; lead_in{label,text}; formula{text}; bullets{items[]}; table{header[],rows[][], template?}; figure{caption,image,explanation?}; verdict{text}. Общий паттерн: любой text может содержать "$unresolved" (линт FAIL). required: meta, sections; additionalProperties:false.
  - **Why:** контракт блоковой грамматики (механизм 2 дизайна).
  - **Inputs:** дизайн §Ключевые механизмы.
  - **Expected output:** валидная JSON Schema draft-07.
  - **Verification:** `python3 -c "import json;json.load(open('.../content-manifest.schema.json'))"` exit 0.

- [ ] **Step 2.2: style-profile.schema.json**
  - **What:** схема: {page:{width_emu,height_emu,orientation,margins_emu}, style_roles:{normal,heading1,heading2,heading3,list_bullet: каждая {font,size_pt,bold,color_hex,spacing_before_emu,spacing_after_emu}}, table_styles:[{name,cell_size_pt,header_bold}], title_block:[{size_pt,bold,italic,alignment}], block_patterns:{figure_caption_regex,formula_style,lead_in_terminator}, microtypography:{decimal_separator,minus_char,percent_spacing,range_dash,quotes,arrow,approx,multiplication}, table_templates:[{name,header[],exemplar_table_index}], section_blueprint:[{level,heading,block_sequence[]}], glossary_terms[]}
  - **Why:** контракт профиля стиля (механизмы 1, 3).
  - **Inputs:** дизайн §Профиль стиля эталона.
  - **Expected output:** валидная JSON Schema draft-07.
  - **Verification:** аналогично 2.1.

## Task 3: База анти-сленга

- [ ] **Step 3.1: base/slang-blacklist.txt**
  - **What:** формат — строки `категория<TAB>regex` (Python re, флаг IGNORECASE). Категории: first_person: `\b(я|мы|нас|нам|наш|наши|вы|вас|вам|ваш)\b`; agent_phrases: `\b(давайте|рассмотрим|я проанализировал|на мой взгляд|к сожалению|итак|по сути|в целом|готово|отлично|супер)\b`; marketing: `\b(инновационн\w+|уникальн\w+|лидер рынка|лучшее решение|революционн\w+)\b`; office_anglicisms: `\b(деливери|фидбек|синк|капля|хайлайт\w*|чекнуть|запушить|мерж\w*|реюз\w*|болванка|скелет отчёта|апдейтнуть)\b`; emoji: `[\U0001F300-\U0001FAFF\u2600-\u27BF]`; markdown_residue: `(\*\*|^#{1,6}\s|\`|\|---)`; exclamations: `!`; placeholders: `(\$unresolved|\bTBD\b|\bTODO\b|\bFIXME\b|\?\?\?)`. Комментарии строками с `#`.
  - **Why:** механизм 5 (анти-сленг), FAIL-категории: все, кроме glossary-warn.
  - **Inputs:** дизайн.
  - **Expected output:** файл с 8 категориями + комментарии.
  - **Verification:** `python3 -c "import re;[re.compile(l.split('\t')[1]) for l in open('...') if l.strip() and not l.startswith('#')]"` exit 0.

## Task 4: extract_profile.py — ядро извлечения

- [ ] **Step 4.1: Интроспекция и style-profile.json**
  - **What:** CLI `python3 extract_profile.py --sample X.docx --outdir work/`. Извлекает: page (из sections[0]); style_roles для Normal/Heading 1/2/3/List Bullet (эффективные font/size/bold/color с учётом наследования стилей; при None в стиле — брать из docDefaults/темы и пометить source); table_styles (имя стиля таблиц + cell_size_pt из первого data-селла + header_bold из первого run header-строки); microtypography (счётчики по всем текстам: `\d\.\d` vs `\d,\d`; U+2212 vs «-» перед цифрами; `\d %` vs `\d%`; `\d–\d`; «» vs ""; →; ≈; ×); glossary_terms (латинские токены ≥2 символов с частотностью, отсортированные по убыванию; ID-паттерны типа POT_\d+; имена файлов \w+\.(py|json|md)). Всё в work/style-profile.json.
  - **Why:** механизм 1 — детерминированное извлечение.
  - **Inputs:** образец docx.
  - **Expected output:** work/style-profile.json, валидный против style-profile.schema.json.
  - **Verification:** на эталоне: normal.font=Calibri, normal.size_pt=11, heading1.color_hex=365F91, heading1.size_pt=14, table_styles содержит "Light Grid Accent 1" с cell_size_pt=10, header_bold=true; microtypography.decimal_separator="point"; glossary содержит per-pot, MAE, fold.

- [ ] **Step 4.2: Детекторы блоковых паттернов**
  - **What:** дополнить profile: title_block (первые Normal-абзацы до первого Heading: для каждого {size_pt,bold,italic,alignment}); block_patterns.figure_caption_regex=r'^Рис\\. \\d+\\.' (проверить по текстам; если не матчит — fallback r'^Рисунок \\d+'); formula_style="italic" (если ≥3 непустых Normal-абзаца полностью курсивные); lead_in_terminator=":" (жирный run в начале абзаца, заканчивающийся на «:» — посчитать встречаемость, ≥2 → паттерн активен).
  - **Why:** механизм 2 — паттерны эталона для блоковой грамматики.
  - **Inputs:** образец docx.
  - **Expected output:** блоки в style-profile.json; для эталона title_block = 3 абзаца [{20,bold},{14,italic},{11}].
  - **Verification:** python-проверка значений title_block в json.

- [ ] **Step 4.3: Кластеризация таблиц → table_templates + section_blueprint**
  - **What:** сигнатура таблицы = header-строка, нормализованная (lowercase, strip, join "|"). Группировать одинаковые сигнатуры; для группы ≥1 таблицы создать шаблон. Имя по эвристике: содержит "mae"|"δ"|"δ"→metrics_comparison; "риск"→risks_register; 2 колонки и первая содержит "параметр"|"период"|"версия"|"framework"→key_value_params; первая колонка "pot"→per_pot_params; "гипотеза"→hypotheses_verdicts; "версия"→version_history; иначе table_template_N. Каждый шаблон: {name, header[], exemplar_table_index, occurrences}. section_blueprint: дерево заголовков (level, heading без номеров, block_sequence — последовательность типов блоков между этим и следующим заголовком: thesis/lead_in/formula/bullets/table/figure по детекторам 4.2 + таблицы).
  - **Why:** механизм 3 — component-library.
  - **Inputs:** образец docx.
  - **Expected output:** в эталоне ≥6 шаблонов, включая metrics_comparison, risks_register, per_pot_params, hypotheses_verdicts; section_blueprint ≈10 разделов H1.
  - **Verification:** python-проверка списка имён шаблонов.

- [ ] **Step 4.4: template.docx**
  - **What:** копия образца: удалить из body ВСЕ дочерние элементы кроме финального sectPr (w:body → оставить w:sectPr). Сохранить work/template.docx. Стили, тема, нумерация, fontTable остаются нетронутыми (они вне body).
  - **Why:** механизм 1 — template-наследование.
  - **Inputs:** образец docx.
  - **Expected output:** work/template.docx открывается python-docx; d.paragraphs пусто; d.tables пусто; d.styles содержит Heading 1 c теми же props; sectPr совпадает с образцом.
  - **Verification:** python-сравнение styles эталона и template: для 5 ролей font/size/color совпадают; margins совпадают.

- [ ] **Step 4.5: Черновики styleguide.draft.md и glossary.draft.md**
  - **What:** styleguide.draft.md: секции Тон (безличность — авто-детект: 0 вхождений 1-го лица в текстах образца; констатация), Структура абзацев (тезис-первое-предложение — пример из образца), Числа и типографика (из microtypography с примерами), Блоки (легенда типов с примерами из образца), Таблицы (шаблоны с заголовками). glossary.draft.md: таблица Термин|Частота|Пример употребления (топ-40). Пометка в шапке каждого: «Черновик сгенерирован автоматически — курайте перед использованием».
  - **Why:** этап 2 конвейера — курирование.
  - **Inputs:** style-profile.json + тексты образца.
  - **Expected output:** 2 md-файла в work/.
  - **Verification:** оба файла непустые, содержат заголовки разделов; grep примера типографики «−12.2 %» в styleguide.

- [ ] **Step 4.6: Копия блэклиста в work/**
  - **What:** скопировать base/slang-blacklist.txt → work/slang-blacklist.txt (пер overriding: если в образце встречаются термины из office_anglicisms — закомментировать эти строки с пометкой «легально в образце»).
  - **Why:** правило «всё из образца легально».
  - **Inputs:** base/slang-blacklist.txt, glossary_terms.
  - **Expected output:** work/slang-blacklist.txt.
  - **Verification:** diff с base; для эталона строка «baseline» НЕ в блэклисте (её там и нет — проверить, что легальные термины образца не пересекаются с категориями).

## Task 5: build_report.py

- [ ] **Step 5.1: Валидация manifest**
  - **What:** CLI `python3 build_report.py --manifest m.json --profile p.json --template t.docx -o out.docx`. Загрузка + jsonschema-валидация (vendored минимальная проверка: required-поля, типы блоков, enumerate level; без внешней jsonschema-библиотеки — ручная проверка в коде). Ошибки — по-русски с jsonpath-подобным путём.
  - **Why:** понятные ошибки для целевой модели.
  - **Inputs:** content-manifest.json.
  - **Expected output:** либо parsed-объект, либо exit 2 с сообщением.
  - **Verification:** manifest без поля meta → exit 2, сообщение содержит "meta".

- [ ] **Step 5.2: Сборка docx**
  - **What:** на базе template.docx (Document(template)): титульный блок из meta по паттерну title_block профиля (центрированные абзацы: title size/bold из title_block[0], subtitle из [1], строка «Проект: … | Версия: … Период: …» из [2]); для каждой section: заголовок стилем Heading N (текст с номером как в manifest.heading); блоки: thesis→Normal; lead_in→Normal с жирным run label+« »; formula→Normal italic; bullets→List Bullet; verdict→List Bullet; table→add_table со стилем из profile.table_styles[0].name, header из template ref или manifest, все ячейки 10pt, header-row bold, шрифт Calibri явно; figure→Normal-абзац caption (текст как есть), следующий абзац — картинка add_picture(width=доступная ширина страницы из профиля), потом Normal explanation. Пустой абзац после title-блока. Сохранить out.docx.
  - **Why:** санкционированный маппинг — дрейф стиля невозможен.
  - **Inputs:** manifest + template + profile.
  - **Expected output:** out.docx.
  - **Verification:** открыть python-docx: стили параграфов соответствуют типам блоков; таблицы в стиле Light Grid Accent 1; у 3+ таблиц header bold; картинки ≥1 в синтетике.

## Task 6: lint_text.py

- [ ] **Step 6.1: Линтер текста (manifest и docx)**
  - **What:** CLI `--manifest m.json [--profile p.json --glossary g.md]` или `--docx d.docx [--blacklist b.txt]`. Извлечь все тексты (из blocks или из параграфов+таблиц docx). Для каждой категории блэклиста найти match'и; placeholders → FAIL; прочие категории → FAIL (по одному сообщению на категорию с цитатой ≤60 символов и номером блока/абзаца). Предупреждения (не FAIL): glossary-warn (латинские токены не из glossary_terms), para_length (абзац >5 предложений). Дополнительно: typography-check против profile.microtypography (десятичная запятая, если профиль «point» → warn с цитатой; ASCII-дефис в отрицательных числах вместо U+2212 → warn; «"» вместо «» → warn). Exit 0 PASS / 1 FAIL. Отчёт на русском в stdout + файл lint_report.txt.
  - **Why:** механизм 5, двухстадийность.
  - **Inputs:** manifest/docx + blacklist.
  - **Expected output:** lint_report.txt + exit code.
  - **Verification:** фикстура с «давайте» и «$unresolved» → exit 1, отчёт содержит обе категории; чистый манифест эталона → exit 0 (после исключения ложных срабатываний: «вы» в словах запрещено границей \b; проверить «Версия» не матчится).

## Task 7: lint_visual.py + render_pdf.py

- [ ] **Step 7.1: lint_visual.py**
  - **What:** CLI `--docx d.docx --profile p.json`. Проверки: каждый абзац использует стиль из допустимого множества (Normal, Heading 1-3, List Bullet — согласно типу контента: заголовки→Heading, bullets-текст→List Bullet, прочее→Normal); эффективные шрифты/кегли ролей соответствуют profile.style_roles (допуск ±0); таблицы используют зарегистрированный table_style, ячейки cell_size_pt, header bold; титульный блок соответствует title_block (размер/жирность/выравнивание); секция страницы/поля совпадают с profile.page. Exit 0/1 + отчёт.
  - **Why:** программная гарантия визуального соответствия.
  - **Inputs:** собранный docx + profile.
  - **Expected output:** отчёт visual_report.txt.
  - **Verification:** на эталоне против его же профиля → exit 0; на docx с «испорченным» стилем (изменить стиль абзаца на Title) → exit 1.

- [ ] **Step 7.2: render_pdf.py**
  - **What:** CLI `--docx d.docx -o dir/`. shutil.which('soffice')/('libreoffice'); если нет — печать предупреждения «PDF недоступен: soffice не найден» и exit 3. Иначе subprocess soffice --headless --convert-to pdf --outdir. Проверка существования pdf. Exit 0.
  - **Why:** двойной выход (docx+pdf).
  - **Inputs:** docx.
  - **Expected output:** d.pdf.
  - **Verification:** в нашей среде (soffice есть) pdf создаётся, размер >0.

## Task 8: SKILL.md ×5 + README

- [ ] **Step 8.1: report-kit-orchestrator/SKILL.md**
  - **What:** frontmatter name/description (рус. триггер: «создай отчёт по образцу» / «новый квартальный отчёт»). Тело: пререквизиты (проверка python3 -c "import docx"), входы (образец.docx, materials/), пошаговый конвейер 0–7 с точными командами всех скриптов, правила циклов правок (lint FAIL → исправить → повторить; max 3 итерации, потом стоп и запрос пользователю), список STOP-условий ($unresolved, отсутствие числа, отсутствие графика). Обязательный раздел «Запрещено»: выдумывать числа, свободный формат, обходить линт.
  - **Why:** входная точка для слабой модели.
  - **Inputs:** спека.
  - **Expected output:** SKILL.md.
  - **Verification:** содержит все 5 команд скриптов (grep).

- [ ] **Step 8.2: report-style-extract/SKILL.md**
  - **What:** триггер (есть образец, нет work/), команда extract_profile.py, что создаётся (5 артефактов), правило курирования черновиков (проверить title_block, glossary, активные паттерны; поправить текст в styleguide.md), что делать при странном образце (нет Heading-стилей → стоп, запрос пользователю).
  - **Verification:** grep команды и списка артефактов.

- [ ] **Step 8.3: report-content-plan/SKILL.md**
  - **What:** триггер (есть materials/ и work/styleguide.md). Правила разворота: mapping outline→section_blueprint; эксперимент = Цель→Метод→Результаты→Интерпретация (4 блока, Результаты = таблица по шаблону + числа с контекстом); график = figure-блок {caption «Рис. N. …», explanation «что изображено → что видно → вывод»}; каждый тезис = 2–5 предложений, безлично, тезис-первым-предложением; термины только из glossary; числа ТОЛЬКО из materials — иначе $unresolved; схема content-manifest; примеры КАЖДОГО типа блока (скопировать из эталона: thesis-пример [7], lead_in «Бизнес-вопрос:», formula [36], figure «Рис. 1»).
  - **Verification:** grep «$unresolved», «Цель», «Рис.»; 7 типов блоков в примерах.

- [ ] **Step 8.4: report-docx-build/SKILL.md**
  - **What:** триггер, команда build_report.py, смысл ошибок валидации (по-русски, с путями), что делать при ошибке (исправить manifest, не править скрипт/шаблон).
  - **Verification:** grep команды.

- [ ] **Step 8.5: report-style-lint/SKILL.md**
  - **What:** триггер, две стадии (manifest → docx+visual), команды, формат отчётов, категории FAIL vs warn, правило цикла правок, стоп после 3 итераций.
  - **Verification:** grep 3 команд (lint_text --manifest, lint_text --docx, lint_visual).

- [ ] **Step 8.6: README.md финал**
  - **What:** заполнить: Назначение, Пререквизиты + команда проверки, Быстрый старт (полный конвейер с командами), Структура, Развёртывание в контуре (скопировать каталог, положить образец и materials, запустить orchestrator), Ограничения (закрытый контур, без MCP; PDF требует soffice).
  - **Verification:** grep «soffice», «Копировать», «python-docx».

## Task 9: Тесты

- [ ] **Step 9.1: tests/test_golden.py — golden-клон**
  - **What:** скрипт-тест (не pytest-unittest, а исполняемый main с exit code): 1) extract_profile.py --sample эталон --outdir tests/work_golden; 2) построить manifest из самого эталона программно (заголовки→sections; абзацы между заголовками → блоки по детекторам: bullets-стиль→bullets{items}, курсивные→formula, lead-in→lead_in, «Рис. N.»→figure{caption, image — сопоставить ближайшую картинку из media эталона}, прочие Normal→thesis; таблицы→table); 3) build_report.py; 4) lint_visual.py → exit 0; 5) программный diff: для каждого параграфа построенного и эталонного (в порядке) — имя стиля совпадает; для таблиц — имя стиля и размерность. Отчёт: N стилевых расхождений = 0.
  - **Why:** главная гарантия — стиль воспроизводится 1-в-1.
  - **Inputs:** эталон.
  - **Expected output:** exit 0, stdout «GOLDEN PASS: 0 style diffs».
  - **Verification:** запуск в нашей среде — PASS.

- [ ] **Step 9.2: tests/test_lint_fixtures.py**
  - **What:** фикстуры: 8 негативных (по 1 на категорию) + 5 позитивных (чистые тексты эталона [5],[7],[51],[103], таблица метрик). Запуск lint_text на каждом; негативные → exit 1 с ожидаемой категорией, позитивные → exit 0.
  - **Verification:** все 13 кейсов проходят с ожидаемыми кодами.

- [ ] **Step 9.3: Синтетический прогон materials→docx+pdf**
  - **What:** tests/materials_synthetic/: outline.md (4 раздела: Аннотация, 1. Контекст, 2. Эксперимент E1, 3. Выводы), draft-заметки (наброски 3-5 строк на раздел), numbers.json (metrics_comparison-строки: Метод A/B, MAE 0.038/0.033), charts/dummy.png (matplotlib scatter 200 точек) + описание. Скрипт tests/test_synthetic.py: полный конвейер (extract по эталону в общую work_synthetic; content-план НЕ моделью, а заготовленным manifest_synthetic.json — проверка конвейера) → lint → build → lint_visual → render_pdf. Exit 0 = все линты зелёные + оба файла существуют.
  - **Why:** end-to-end проверка двойного выхода.
  - **Inputs:** эталон + синтетика.
  - **Expected output:** tests/out_synthetic/report.docx + report.pdf.
  - **Verification:** запуск — exit 0, `ls` обоих файлов.

## Task 10: Финальный аудит

- [ ] **Step 10.1: Copy-out аудит**
  - **What:** grep по всем файлам комплекта: нет «/root/» (кроме комментариев-примеров — заменить на относительные), нет «lite_rag|skills_kb|tool-discovery|MCP», нет «pip install» без пометки «выполнено при подготовке контура». python3 -m py_compile на всех скриптах.
  - **Why:** автономность закрытого контура.
  - **Inputs:** весь комплект.
  - **Expected output:** аудит-отчёт в stdout.
  - **Verification:** grep-команды возвращают пусто.

- [ ] **Step 10.2: Итоговая сводка**
  - **What:** вывести ls -R комплект + сводку: сколько файлов, размеры, результаты 3 тестов.
  - **Verification:** сводка содержит PASS по всем трём тестам.

## Verification gates

- После Task 4: показать пользователю извлечённый profile эталона (ключевые значения).
- После Task 9: показать результаты golden/slang/synthetic.
- После Task 10: комплект готов к копированию в контур.
