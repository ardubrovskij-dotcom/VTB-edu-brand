# Макеты и формат deck.json

## Структура файла

```json
{
  "title": "Название презентации (метаданные PDF/PPTX)",
  "kicker": "Кикер по умолчанию — левый верх каждого слайда",
  "footer": "Подвал по умолчанию — левый низ",
  "slides": [ { "layout": "cover", "...": "..." } ]
}
```

Общие поля любого слайда: `kicker`, `footer` (переопределяют общие), `notes` (заметки докладчика → PPTX),
`hide_page` (скрыть номер). В текстовых полях допустим инлайн-HTML: `<b>`, `<br>`, `&nbsp;`.
Пути к картинкам — относительно deck.json.

Поля, общие для светлых макетов: `title`, `lead`, `thesis` + `thesis_label` (градиентная карточка внизу),
`source` (строка «Источник: …» внизу).

## Каталог

| layout | Фон | Обязательные поля | Необязательные | Лимиты, чтобы влезло |
|---|---|---|---|---|
| `cover` | тёмный | `title` | `subtitle`, `meta[]` (чипы) | заголовок ≤ 70 знаков, 2–4 чипа |
| `section` | тёмный | `title` | `num` ("01"), `subtitle` | заголовок ≤ 45 знаков |
| `agenda` | светлый | `title`, `items[]` (строка или `{title,text}`) | — | 2–8 пунктов |
| `text` | светлый | `title` | `lead`, `text` (строка или массив), `thesis`, `source` | ≤ 90 слов на слайд |
| `bullets` | светлый | `title`, `items[]` | `lead`, `side{label,text}`, `thesis`, `source` | 3–6 пунктов по ≤ 12 слов |
| `cards` | светлый (`dark: true` — тёмный) | `title`, `items[{title,text,num?,highlight?}]` | `columns`, `numbered:false`, `thesis` | 2–4 в ряд (6 = 3×2), текст ≤ 25 слов |
| `stats` | светлый (`dark: true`) | `title`, `items[{value,unit?,label}]` | `lead`, `thesis`, `source` | 2–4 числа, `value` ≤ 5 знаков |
| `process` | светлый | `title`, `items[{title,text}]` | `highlight_last:false`, `thesis` | 3–5 шагов, текст ≤ 15 слов |
| `compare` | светлый | `title`, `columns[2]{title,items[],highlight?}` | `lead`, `thesis` | ≤ 6 пунктов в колонке |
| `table` | светлый | `title`, `headers[]`, `rows[][]` | `align[]` ("right" для чисел), `highlight[[r,c]]`, `source` | ≤ 7 строк × 5 колонок |
| `chart` | светлый | `title`, `items[{label,value,highlight?,display?}]` | `max`, `unit`, `lead`, `source` | 3–8 столбцов |
| `image` | светлый | `title`, `image` | `image_side` (left/right), `text`, `items[]`, `thesis` | ≤ 50 слов |
| `person` | светлый | `name` | `photo`, `role`, `label` ("Лектор"), `text`, `items[]` | фото ≥ 520×700 px |
| `quote` | тёмный | `text` | `author`, `role`, `photo` | ≤ 30 слов |
| `closing` | тёмный | — | `title` ("Спасибо!"), `subtitle`, `button`, `contacts[{label,value}]` | ≤ 4 контакта |

## Как выбирать макет

| Содержание | Макет |
|---|---|
| Начало / конец | `cover` / `closing` |
| Переход между частями (> 8 слайдов) | `section` |
| План встречи, программа | `agenda` |
| Одна мысль + пояснение | `text` с `thesis` |
| Перечень без иерархии | `bullets` (с `side`, если есть вывод) |
| 2–4 равноправных объекта | `cards` |
| Ключевые числа | `stats` (с `source` обязательно) |
| Последовательность, алгоритм | `process` |
| До/после, вариант А/Б | `compare` |
| Точные значения по нескольким признакам | `table` |
| Сравнение величин одного ряда | `chart` |
| Лектор, эксперт | `person` |
| Сильная формулировка | `quote` |

## Отчёт о переполнении

`report.json` и вывод скрипта: `ПЕРЕПОЛНЕНИЕ` — содержимое вышло за рабочую область (top 168 … bottom 128).
Код выхода 2. Исправление: сократить текст, разбить слайд на два, уменьшить число карточек/строк.
Уменьшать кегль вручную нельзя — шкала зафиксирована в `brand.md`.
