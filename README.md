# VTB-edu-brand

Скиллы Claude для материалов в фирменном стиле «ВТБ Образование».

| Скилл | Что делает |
|---|---|
| [`vtb-edu-presentations`](vtb-edu-presentations/SKILL.md) | Редактируемые презентации PPTX, собранные из настоящего мастер-шаблона «мерико × ВТБ Образование»: 38 макетов, 8 направлений (мастеров), фоны, карточки, портреты, векторные иконки и иллюстрации шаблона. |

## Быстрый старт

```bash
pip install python-pptx pillow lxml
python3 vtb-edu-presentations/scripts/build_deck.py vtb-edu-presentations/examples/example.json -o out/ --preview
```
`--preview` дополнительно делает PDF и PNG (нужны LibreOffice и poppler-utils).

## Установка как скилла
- **Claude Code**: скопировать папку `vtb-edu-presentations` в `~/.claude/skills/` (или в `.claude/skills/` проекта).
- **claude.ai**: заархивировать папку `vtb-edu-presentations` в zip и загрузить в настройках claude.ai (раздел со скиллами).

## Источник бренда
Папка «ВТБ О брендинг» (Google Drive): шаблон презентации PPTX, «Фоны на слайды», «Карточки для презентации», «Люди».
Что откуда взято и как обновить — [`references/brand.md`](vtb-edu-presentations/references/brand.md),
[`references/assets.md`](vtb-edu-presentations/references/assets.md).
