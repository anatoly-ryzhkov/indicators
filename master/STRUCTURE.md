# Структура системы оценки

```
ЗАДАЧА РОЛИ
    │
    ├── task_map.yaml                  ← группы задач → задачи → критерии + вопросы
    │       task_group                 (TG-01, TG-02, ...)
    │         └── task                 (T-01, T-02, ...)
    │               ├── criteria_fit   [C01: fit, C05: partial, ...]
    │               └── questions      [{text, criterion, type}, ...]
    │
    └── indicators_full_v2.xlsx        ← критерии → индикаторы
            Indicators sheet
              criterion_id → [C01-L0-F1-i01, C01-L0-F2-i02, ...]
              level, blocks_level_up_to, indicator_en, ...

КЕЙС (cases/case_XX/)
    │
    ├── meta.yaml                      ← criteria_in_scope, master_version, ...
    ├── placeholder_map.yaml           ← {PLACEHOLDER} → значения
    ├── adapted_task_map.yaml          ← task_map с подставленными плейсхолдерами + реальные вопросы
    ├── adapted_indicators.xlsx        ← indicators с подставленными плейсхолдерами
    ├── human_grades.yaml              ← {C01: -1, C08: 0, ...}
    └── validation/
            run_output.json            ← LLM результаты по индикаторам
            diff_report.yaml           ← сравнение pipeline vs human grades

ПОЛНАЯ ЦЕПОЧКА:
    adapted_task_map.yaml
        → вопросы задаются кандидату → транскрипт
        → adapted_indicators.xlsx (индикаторы per criterion)
        → run_output.json (LLM: fired/not fired per indicator)
        → aggregation (validate_case.py) → pipeline grades per criterion
        → diff_report.yaml (pipeline grades vs human_grades)
        → диагностика → фикс master + adapted
```

## Файлы

| Файл | Слой | Шаблон | Кейс |
|------|------|--------|------|
| `master/task_map.yaml` | задачи → критерии + вопросы | ✓ master | — |
| `schemas/task_map.yaml.template` | пустой шаблон | ✓ | — |
| `cases/XX/adapted_task_map.yaml` | задачи + реальные вопросы | — | ✓ кейс |
| `master/indicators_full_v2.xlsx` | критерии → индикаторы | ✓ master | — |
| `cases/XX/adapted_indicators.xlsx` | критерии → индикаторы (адапт.) | — | ✓ кейс |
