# `ru-academic-gost` — Autonomous Russian Academic Paper & GOST Pipeline Skill

Полноценный автономный Agent Skill для **Google Antigravity** (а также совместимых агентов — Claude Code, Cursor, Codex, Gemini CLI), который ведёт студенческую работу российских вузов (**доклад, реферат, эссе, курсовая работа, ВКР / диплом**) от интервью и скачивания реальных русскоязычных PDF до ГОСТ-вёрстки в Microsoft Word и визуального аудита по PNG-превью.

## Что внутри

- **7 строгих фаз с Exit Gates** ([`SKILL.md`](SKILL.md)):
  - **Фаза 0:** Структурированное интервью (или выбор лучшей темы из списка кафедры), генерация `project_contract.json`, `meta.json`, `project_state.json`.
  - **Фаза 1:** Поиск, скачивание и валидация реальных PDF/веб-источников (КиберЛенинка, ВАК/РИНЦ, Росстат, ВЦИОМ, Mediascope, Sostav, AdIndex, VC) + сохранение библиографических `.meta.json`.
  - **Фаза 2:** Архитектура работы, выбор практического кейса + эталонного контр-кейса, бюджет знаков с пробелами (`00_plan.md`).
  - **Фаза 3:** Пошаговое написание разделов строго по извлечённым файлам с прямыми цитатами «ёлочками» (`«...»`), нативными сносками или ссылками `[N, с. X]`.
  - **Фаза 4:** Двухпроходная редактура «Анти-слоп» + детерминированная проверка через [`Vladimir-Human/humanizer-ru`](https://github.com/Vladimir-Human/humanizer-ru) (`humanizer-markers`, `humanizer-scan --genre academic`, **машинный Факт-замок** `humanizer-facts diff`).
  - **Фаза 5:** ГОСТ-вёрстка `.docx` через Pandoc + in-place `python-docx` (титульный лист, живое автообновляемое Содержание Word, нативные надстрочные подстрочные сноски TNR 10 пт, ГОСТ-таблицы с `tblHeader`/`cantSplit`, нумерация со 2-й страницы).
  - **Фаза 6:** 9-пунктовый авто-аудит `.docx` + PDF-экспорт (Microsoft Word COM / LibreOffice) + рендер PNG-превью через PyMuPDF и цикл самокоррекции до `PASS`.
  - **Фаза 7:** Выдача `output/final.docx`, устной речи `output/speech.md` и шпаргалки к защите `output/defense_cheatsheet.md`.

### Скрипты (`scripts/`)

| Скрипт | Назначение |
|---|---|
| `scripts/preflight.py` | Диагностика Python-пакетов, `humanizer-ru`, `pandoc`, Word COM / LibreOffice и PyMuPDF |
| `scripts/fetch_sources.py` | Скачивание PDF (КиберЛенинка, `citation_pdf_url`, прямые PDF) и очистка веб-статей с валидацией кириллического текстового слоя |
| `scripts/parse_pdfs.py` | Колоночный парсинг PDF, расклейка переносов, определение журнальных страниц, MD5/текстовая дедупликация, Exit Gate 1 и `--search` |
| `scripts/analyze_draft.py` | Проверка бюджета знаков, каскадов `-ение/-ание/-ость`, бан-листа клише, верификация цитат `«...»` по PDF, интеграция `humanizer-ru` |
| `scripts/build_gost.py` | Сборка ГОСТ `.docx` (Pandoc + in-place OOXML + живое обновление оглавления в Word) |
| `scripts/audit_docx.py` | 9-пунктовый аудит: заглушки, скрытый Unicode, гомоглифы, пунктуация сносок, библиография (НПА + А–Я с `Ё`=`Е`), страницы, висячие хвосты, `Там же.` вверху страницы, PNG-превью + `--fix` |

### Справочники и шаблоны

- `references/schemas.md` — JSON-схемы контракта, титульного листа и состояния проекта
- `references/search-strategy.md` — стратегия поиска открытых русскоязычных источников
- `references/ru-style.md` — академический стиль, бан-лист штампов, 5-осевая самооценка
- `references/prompts-library.md` — 8 проверенных промптов пайплайна
- `references/gost-rules.md` — правила ГОСТ 7.32-2017 / ГОСТ Р 7.0.5 / ГОСТ Р 7.0.100
- `templates/` — готовые шаблоны `contract_kursovaya.json`, `contract_report.json`, `meta_example.json`

---

## Быстрая установка

### 1. Склонируйте скилл в глобальный каталог навыков

**Google Antigravity (Windows PowerShell):**
```powershell
git clone https://github.com/Mra3ota/ru-academic-gost.git "$HOME\.gemini\config\skills\ru-academic-gost"
```

**Google Antigravity (macOS / Linux):**
```bash
git clone https://github.com/Mra3ota/ru-academic-gost.git ~/.gemini/config/skills/ru-academic-gost
```

### 2. Установите зависимости Python

```bash
python -m pip install --upgrade -r ~/.gemini/config/skills/ru-academic-gost/requirements.txt
```

> На Windows при наличии Microsoft Word пакет `pywin32` обеспечивает автоматическое обновление оглавления (TOC) и точный PDF-рендер без LibreOffice. На macOS/Linux установите LibreOffice (`soffice`) для экспорта PDF.

### 3. Проверьте окружение

```bash
python ~/.gemini/config/skills/ru-academic-gost/scripts/preflight.py
```

---

## Как запускать

Откройте пустую папку под новую работу в **Google Antigravity** и напишите:

- **Курсовая / ВКР:**
  > *«Запусти скилл `ru-academic-gost` в режиме курсовой. Вот методичка и список тем кафедры — сначала задай вопросы и помоги выбрать тему, затем скачай источники и проведи работу до готового `.docx`.»*
- **Доклад / реферат:**
  > *«Запусти скилл `ru-academic-gost` в режиме доклада на 6 страниц + устная речь на 5 минут по теме [...].»*

## Лицензия

[MIT](LICENSE)
