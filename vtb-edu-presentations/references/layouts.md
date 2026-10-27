# Макеты и формат deck.json

```json
{
  "title": "Название файла (метаданные PPTX)",
  "direction": "vtb",
  "repeat_title": true,
  "slides": [ { "layout": "title", "...": "..." } ]
}
```

- `direction` — палитра, фон и портрет по умолчанию: `vtb | family | prof | china | b2b | infra | private |
  privilege` (`brand.md`).
- `repeat_title: true` — последний слайд повторяет титул (так заканчиваются презентации ВТБ Образования).

## Шапка контентного слайда (общая для всех макетов, кроме `title`, `section`)

| Поле | Что это | Вид |
|---|---|---|
| `kicker` | рубрика над заголовком: «Базовый формат:», «ШОиФД-1, осень 2023» | 32 pt, серый, прописными |
| `title` | **вывод слайда целиком**, 1–4 строки | 64 pt, прописными |
| `title_grey` | серая часть заголовка | 64 pt, серый |
| `subtitle` (= `lead`) | уточнение, контекст, цифра | 32 pt, серо-синий |
| `footnote`, `source` | сноска и источник | 20 pt внизу |
| `notes` | заметки докладчика | — |
| `logo: false` | убрать логотип | — |

Контент всех макетов **прижат к низу** (y = 976), между шапкой и контентом остаётся поле. Подробно — `house-style.md`.

Разметка в тексте: `<b>…</b>` — акцентный цвет, `<g>…</g>` — серый, `<s>…</s>` — тёмный полужирный (термин),
`<k>…</k>` — сохранить регистр в прописном заголовке (`<k>ШОиФД</k>`), `\n` — новая строка.
Неразрывные пробелы после коротких слов и перед тире скрипт ставит сам.

Ресурсы — **по имени** из `references/assets.md` (`"photo": "speaker"`, `"icon": "bank"`, `"portrait": "vtb-2"`)
или **путём к своему файлу** относительно deck.json (`"photo": "people/ivanov.png"`).

## Основные макеты (стиль по умолчанию)

| layout | Для чего | Поля | Сколько влезает |
|---|---|---|---|
| `title` | Титул: белое поле + диагональная панель с портретом | `kicker` (название программы, 40 pt серый, можно `\n`), `title` (82 pt), `subtitle` + `date` (подпись внизу), `portrait` (по умолчанию `vtb-2`), `partner_logo`, `background` | заголовок 2–3 строки |
| `section` | Раздел («Приложение») | `kicker`, `title`, `subtitle`, `date`, `num`; `photo` — вариант с фото | 1–2 строки |
| `cards` | 2–4 простые серые карточки | `items[{title, text \| items[], highlight?}]`, `height` | текст до ~30 слов в карточке |
| `stats` | Крупные цифры | `items[{value, unit?, after?, text}]` (`after` — второе крупное число: «80 из 89»), `columns` (3), `height` (259; для одной строки — до 390) | 2–6 цифр, значение до 7 знаков |
| `speakers` | Спикеры, эксперты | `items[{name, role, photo?}]`, `columns` | 2–9 человек; имя в одну строку |
| `features` | Компетенции, преимущества | `card{title, icon}` (левая градиентная), `items[]` (строки), `columns` | 3–6 ячеек по ≤ 25 слов |
| `rows` | Состав предложения, форматы, условия | `rows[{label, icon, items[] \| text, columns?}]` (колонок 1–2), `columns` (по умолчанию 2) | 2–4 строки по 2–6 пунктов |
| `table` | Шаги, сроки, сравнение | `headers[]`, `rows[][]` (`\n` во второй строке ячейки — мелкая подпись), `widths[]`, `align[]`, `highlight_rows[]` | до 9 строк |
| `chart` | Один график без карточки | `charts[{title, type, categories[], series[{name, values[]}], min, max, step, format}]`, `source` | 3–8 категорий, до 4 рядов |
| `course` | Страница курса в приложении | `kicker`, `title`, `description`, `topics[{term, text}]`, `audience{title?, text, result?}` | до 10 тем; описание до 4 строк |
| `contacts` | Контакты | `people[{name, role, email, phone?, photo?}]` | 1–3 человека |

## Дополнительные макеты (по мастер-шаблону)

| layout | Стр. шаблона | Поля |
|---|---|---|
| `toc` | 15 | `items[{title, page?}]` |
| `bullets` | 16–18 | `items[]` (строка или `{text, icon}`), `numbered: false`, `conclusion` |
| `process` | 25 | `items[{title, text, label?}]`, `result: false` |
| `highlights` | 29–31 | `items[{value \| icon, unit?, text, highlight?}]`, `columns` |
| `photo_text` | 34, 35 | `photo`, `items[]`, `label`, `photo_side: left` |
| `menu_photo` | 28 | `photo`, `items[{title, text, icon?, highlight?}]` |
| `timeline` | 48 | `items[{label, text}]` |
| `scheme` | 39 | `top[{title, text}]`, `items[{title, text, num? \| icon?, highlight?}]`, `bottom` |
| `image` | 62 | `photo` |

Варианты старых стилей: `title` с `variant: pattern | split`; `cards` с `style: header` (шапки с иконками);
`table` с `style: blocks`; `chart` с `style: cards` или с двумя графиками.

## Как выбрать макет

| Задача слайда | Макет |
|---|---|
| Открыть и закрыть презентацию | `title` + `repeat_title: true` |
| Почему это важно, 3 составляющие | `cards` |
| Масштаб, результаты, оценки | `stats` |
| Люди | `speakers` |
| Что умеет команда | `features` |
| Что входит в предложение, варианты формата | `rows` (+ `kicker` с названием формата) |
| Данные в поддержку | `chart` |
| План, сроки, ответственные | `table` |
| Блок приложений | `section` |
| Описание курса или программы | `course` |
| Кому писать | `contacts` |

## Проверка текста
Переносы оцениваются по метрикам Open Sans (для полужирного — Open Sans SemiBold). Если текст не помещается,
кегль снижается до минимума, и об этом появляется запись в `report.json`. Если блоки не помещаются по высоте,
пишется предупреждение, а при `ПЕРЕПОЛНЕНИЕ` код выхода — 2. Тогда сократить текст или разбить слайд.
