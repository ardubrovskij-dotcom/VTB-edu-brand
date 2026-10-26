# VTB-edu-brand

Скиллы Claude для материалов в фирменном стиле «ВТБ Образование».

| Скилл | Что делает |
|---|---|
| [`vtb-edu-presentations`](vtb-edu-presentations/SKILL.md) | Редактируемые презентации PPTX по мастер-шаблону «мерико × ВТБ Образование, 2026»: 17 макетов, 8 направлений, фирменные фоны, портреты, фото и иконки. |

## Быстрый старт

```bash
pip install python-pptx pillow
python3 vtb-edu-presentations/scripts/build_pptx.py vtb-edu-presentations/examples/temp_course.json -o out/ --preview
```
`--preview` дополнительно делает PDF и PNG (нужны LibreOffice и poppler-utils).

## Установка как скилла
- **Claude Code**: скопировать папку `vtb-edu-presentations` в `~/.claude/skills/` (или в `.claude/skills/` проекта).
- **claude.ai**: заархивировать папку `vtb-edu-presentations` в zip и загрузить в настройках claude.ai (раздел со скиллами).

## Источник бренда
Папка «ВТБ О брендинг»: «Архитектура визуальной системы», мастер-шаблон презентации, шаблоны и паттерны.
Что откуда взято — [`references/brand.md`](vtb-edu-presentations/references/brand.md).
