# VTB-edu-brand

Скиллы Claude для материалов в фирменном стиле «ВТБ Образование».

| Скилл | Что делает |
|---|---|
| [`vtb-edu-presentations`](vtb-edu-presentations/SKILL.md) | Презентации 16:9 из `deck.json` → PDF, PNG, PPTX, HTML. 15 макетов, проверка переполнения. |

## Быстрый старт

```bash
pip install python-pptx playwright && playwright install chromium   # один раз
python3 vtb-edu-presentations/scripts/build_deck.py vtb-edu-presentations/examples/temp_course.json -o out/
```

## Установка как скилла
- **Claude Code**: скопировать папку `vtb-edu-presentations` в `~/.claude/skills/` (или в `.claude/skills/` проекта).
- **claude.ai**: заархивировать папку `vtb-edu-presentations` в zip и загрузить в настройках claude.ai (раздел со скиллами).

## Источники бренда
См. [`vtb-edu-presentations/references/brand.md`](vtb-edu-presentations/references/brand.md): таблица источников и
открытые вопросы, которые нужно сверить с Figma-файлом «мерико × ВТБ Образование».
