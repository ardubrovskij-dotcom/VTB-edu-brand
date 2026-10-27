# Ресурсы скилла

| Путь | Что | Как указать в deck.json |
|---|---|---|
| `assets/template.pptx` | облегчённый мастер-шаблон: 8 мастеров, слайды-прототипы 9–49 и 52, библиотека иконок и иллюстраций (63–75) | — (берётся сам) |
| `assets/backgrounds/NN_MM.jpg` | 40 фонов (`backgrounds.jpg`) | `"background": "01_03"` в `section` |
| `assets/cards/` | подложки карточек `card_MM`, `highlight_MM` (`cards.jpg`) | — (по направлению) |
| `assets/people/` | портреты: `vtb`, `family`, `prof`, `china`, `b2b`, `infra`, `private`, `privilege`, `extra` (`people.jpg`) | `"portrait": "prof"` |
| `assets/photos/` | фото сцен: `lecture`, `presenter`, `speaker`, `laptop`, `team` (вертикальное), `students` (вертикальное), `group`, `audience-cutout` (PNG без фона) | `"photo": "lecture"` |
| `assets/logo/` | логотип: `logo-color`, `logo-white`, `logo-black`, `logo-onblue` | — |
| `assets/fonts/` | Open Sans (OFL) — только для оценки переносов, если VTB Group не установлен | — |
| `references/icons.png` | 43 иконки (вариант «Заливка» — суффикс `-fill`) | `"icon": "bank"` |
| `references/illustrations.png` | 114 иллюстраций | `"icon": "ill-finance-05"` |

Список имён иконок и иллюстраций без картинки: `python3 scripts/build_deck.py --list-icons`.

## Подбор иконок по смыслу
- обучение: `graduation`, `presentation`, `bulb`, `search`, `timer`, `edit`, `sign`
- финансы: `bank`, `money`, `coins`, `card`, `calculator`, `briefcase`
- ИТ: `monitor`, `server`, `cloud`, `link`, `globe`, `folder-in`
- люди: `people`, `people-list`, `user-add`, `user-growth`, `user-target`, `user-chat`, `handshake`
- управление: `structure`, `layers`, `folder`, `image`, `list-fill`, `dots`
- связь: `mail`, `send`, `phone`, `chat`
- знаки: `badge`, `bell`, `pin`, `home`, `copy`, `download`, `star`, `gear`

## Обновление шаблона
Если маркетинг обновил шаблон или папку «ВТБ О брендинг»:
1. Скачать новый PPTX и запустить
   `python3 tools/prepare_template.py "<шаблон>.pptx" -o assets/template.pptx`.
   Скрипт оставит слайды 9–49, 52, 63–75 (по номерам исходного шаблона), удалит лишние макеты, пережмёт картинки
   (58 → 14 МБ) и уберёт встроенный шрифт VTB Group (`--keep-fonts` — оставить).
2. Если номера или состав слайдов-образцов изменились — поправить `proto` и номера плейсхолдеров в
   `scripts/layouts.py`, имена иконок — в `ICON_ROWS` в `scripts/build_deck.py`.
3. Фоны, карточки и портреты — заменить файлы в `assets/` с теми же именами.
4. Собрать `examples/example.json` с `--preview` и проверить картинки.

## Почему шрифт не встроен
Репозиторий публичный, а VTB Group — корпоративный шрифт. В PPTX прописаны `VTB Group Book` / `VTB Group Light`:
на компьютерах ВТБ со шрифтом презентация выглядит как шаблон, без шрифта PowerPoint подставит замену.
