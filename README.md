# VTB-edu-brand

Плагин Claude со скиллами в фирменном стиле «ВТБ Образование».

| Скилл | Что делает |
|---|---|
| [`vtb-edu-presentations`](skills/vtb-edu-presentations/SKILL.md) | Редактируемые презентации PPTX, собранные из мастер-шаблона «мерико × ВТБ Образование»: 38 макетов, 8 направлений, фоны, карточки, 9 портретов с именами, векторные иконки и иллюстрации шаблона. |

## Установка

### Claude Code (рекомендуется)
```
/plugin marketplace add ardubrovskij-dotcom/VTB-edu-brand
/plugin install vtb-edu-brand@vtb-edu-brand
```
На Claude Code 2.1.275 и новее — одной командой:
```
/plugin install vtb-edu-brand --marketplace ardubrovskij-dotcom/VTB-edu-brand
```
Обновление после изменений в репозитории: `/plugin marketplace update vtb-edu-brand`.

Без плагина: скопировать папку `skills/vtb-edu-presentations` в `~/.claude/skills/` (или в `.claude/skills/` проекта).

### claude.ai
Скачать `vtb-edu-presentations.zip` из релиза [latest](https://github.com/ardubrovskij-dotcom/VTB-edu-brand/releases/tag/latest)
и загрузить в claude.ai → Настройки → Возможности → Скиллы. Zip собирается автоматически
(`.github/workflows/skill-zip.yml`) при каждом изменении в `skills/`. Вручную:
```bash
cd skills && zip -r ../vtb-edu-presentations.zip vtb-edu-presentations -x '*/__pycache__/*'
```

## Быстрый старт без Claude
```bash
pip install python-pptx pillow lxml
python3 skills/vtb-edu-presentations/scripts/build_deck.py skills/vtb-edu-presentations/examples/example.json -o out/ --preview
```
`--preview` дополнительно делает PDF и PNG (нужны LibreOffice и poppler-utils).

## Структура
```
.claude-plugin/          манифест плагина и маркетплейса, иконка
skills/vtb-edu-presentations/
  SKILL.md               инструкция для Claude
  scripts/               build_deck.py (сборка), layouts.py (макеты, направления, портреты)
  tools/                 prepare_template.py (обновление шаблона)
  assets/                облегчённый шаблон, фоны, карточки, портреты, фото, логотип
  references/            описание макетов, бренда, ресурсов; превью
  examples/              пример колоды и превью
```

## Источник бренда
Папка «ВТБ О брендинг» (Google Drive): шаблон презентации PPTX, «Фоны на слайды», «Карточки для презентации», «Люди».
Что откуда взято и как обновить — [`references/brand.md`](skills/vtb-edu-presentations/references/brand.md),
[`references/assets.md`](skills/vtb-edu-presentations/references/assets.md).
