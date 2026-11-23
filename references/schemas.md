# Схемы контрактов и артефактов проекта (`project_contract.json`, `meta.json`, `project_state.json`)

Используй эти точные ключи JSON — все скрипты (`fetch_sources.py`, `parse_pdfs.py`, `analyze_draft.py`, `build_gost.py`, `audit_docx.py`) опираются именно на них.

---

## 1. `project_contract.json` (создаётся на Exit Gate 0)

```json
{
  "work_type": "kursovaya",
  "autonomy_mode": "standard",
  "topic": "Точная формулировка темы без точки в конце",
  "discipline": "Название дисциплины",
  "pages": {
    "min": 20,
    "max": 25,
    "target": 23
  },
  "oral_speech": {
    "enabled": false,
    "duration_minutes": 5,
    "words_per_minute": 125
  },
  "citations": {
    "style": "footnote",
    "forbid_tam_zhe_page_top": true,
    "min_quotes_per_theory_page": 1,
    "target_quote_share_pct": [10, 20]
  },
  "bibliography": {
    "min_sources": 20,
    "max_sources": 25,
    "min_year": 2020,
    "max_classic_sources": 3,
    "forbid_untranslated_foreign": true,
    "npa_first": true,
    "sort_order": "alphabetical_ru"
  },
  "layout": {
    "font_name": "Times New Roman",
    "font_size_pt": 14,
    "line_spacing": 1.5,
    "first_line_indent_cm": 1.25,
    "margins_mm": {
      "left": 30,
      "right": 15,
      "top": 20,
      "bottom": 20
    },
    "footnote_font_size_pt": 10,
    "table_font_size_pt": 12,
    "include_title_page": true,
    "include_toc": true,
    "page_number_position": "bottom_center",
    "min_lines_last_section_page": 4,
    "short_table_max_rows": 5
  },
  "conclusion_style": "analytical_summary",
  "student_voice_sample": "",
  "special_instructions": []
}
```

### Допустимые значения ключевых полей
- `work_type`: `"report"` (доклад), `"referat"` (реферат), `"essay"` (эссе), `"kursovaya"` (курсовая), `"vkr"` (диплом/ВКР).
  - Для `report` / `essay` по умолчанию: `include_toc: false` (если объём ≤ 10 стр.), `min_sources: 5..10`.
  - Для `kursovaya`: `include_toc: true`, `min_sources: 20..25`.
  - Для `vkr`: `include_toc: true`, `min_sources: 40..60`, на титульнике опционально включаются поля рецензента/допуска.
- `autonomy_mode`: `"standard"` (1 пауза после Фазы 2) | `"autopilot"` (без пауз до `final.docx`).
- `citations.style`:
  - `"footnote"` — постраничные подстрочные сноски Word (Markdown-синтаксис `[^id]` + `[^id]: ...`).
  - `"brackets"` — внутритекстовые/затекстовые ссылки `[N, с. X]` (синхронизируются с нумерацией `05_biblio.md`).
- `layout.page_number_position`: `"bottom_center"` (стандарт ГОСТ 7.32-2017) | `"bottom_right"` | `"top_center"`.
- `conclusion_style`: `"strict_chapter_findings"` | `"analytical_summary"` | `"recommendations"`.

---

## 2. `meta.json` (данные титульного листа)

Опциональные пустые поля (`""`) пропускаются сборщиком `build_gost.py` автоматически. Не оставляй заглушек вида `____`, `[...]`, `ХХ-000` или `TODO`.

```json
{
  "ministry": "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
  "university": "Федеральное государственное бюджетное образовательное учреждение высшего образования «Название вуза»",
  "institute": "Институт / Факультет",
  "department": "Кафедра ...",
  "work_label": "КУРСОВАЯ РАБОТА",
  "discipline": "по дисциплине «...»",
  "topic": "Полная формулировка темы без точки в конце",
  "direction": "Направление подготовки 00.00.00 «...»",
  "profile": "",
  "student_name": "Фамилия Имя Отчество",
  "student_group": "группа XX-000",
  "course": "3 курса очной формы обучения",
  "supervisor_role": "Научный руководитель:",
  "supervisor_title": "к.э.н., доцент",
  "supervisor_name": "Иванов И. И.",
  "reviewer_title": "",
  "reviewer_name": "",
  "city": "Москва",
  "year": "2026"
}
```

---

## 3. `project_state.json` (контроль прогресса и возобновление)

Обновляй `project_state.json` после завершения каждой фазы — это позволяет безопасно продолжать длинную ВКР/курсовую после `/compact` или нового диалога.

```json
{
  "current_phase": 3,
  "completed_phases": [0, 1, 2],
  "sources_valid": 23,
  "drafts": {
    "01_intro.md": {"target_chars": 3200, "actual_chars": 3180, "humanized": true, "score_50": 41},
    "02_ch1.md": {"target_chars": 11500, "actual_chars": 11420, "humanized": false, "score_50": null}
  },
  "audit_iterations": 0,
  "last_audit_status": null
}
```

---

## 4. Соглашение об именовании черновиков в `drafts/`

`build_gost.py` собирает файлы в `drafts/` строго в алфавитном порядке, игнорируя `00_plan.md`, `sources_*` и подпапки (`extracted/`, `snapshots/`):
- `01_intro.md` — `# ВВЕДЕНИЕ` (или `# Введение`)
- `02_ch1.md` — `# ГЛАВА 1. ...` (или `# 1 ...` для доклада)
- `03_ch2.md` — `# ГЛАВА 2. ...`
- `04_ch3.md` — `# ГЛАВА 3. ...` (если есть 3-я глава, например в ВКР)
- `05_conclusion.md` (или `04_conclusion.md`) — `# ЗАКЛЮЧЕНИЕ`
- `06_biblio.md` (или `05_biblio.md`) — `# СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ`
- `07_appendix_a.md` (опционально) — `# ПРИЛОЖЕНИЕ А`
