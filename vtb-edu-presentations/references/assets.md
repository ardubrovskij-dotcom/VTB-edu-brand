# Фирменные ресурсы

Все файлы — из материалов «мерико × ВТБ Образование, 2026». В deck.json ресурсы указываются по имени без расширения.

## Иконки — `"icon": "<имя>"`

Линейные, Retail Blue; для цветных плашек скрипт сам берёт белую версию `<имя>-white.png`. Наглядно: `references/icons.png`.

`bank`, `bell`, `briefcase`, `bulb`, `calculator`, `card`, `chat`, `check`, `cloud`, `coins`, `copy`, `document`, `download`, `edit`, `edit-2`, `folder`, `folder-in`, `gear`, `globe`, `globe-2`, `graduation`, `handshake`, `home`, `image`, `layers`, `link`, `list`, `mail`, `mail-2`, `money`, `monitor`, `people`, `people-2`, `people-list`, `phone`, `pin`, `presentation`, `search`, `send`, `send-2`, `server`, `sign`, `signature`, `star`, `structure`, `timer`, `user-add`, `user-chat`, `user-growth`, `user-target`

Подбор по смыслу: обучение — `graduation`, `presentation`, `bulb`, `search`, `timer`; финансы — `bank`, `calculator`, `money`, `coins`, `card`, `briefcase`, `document`; ИТ — `monitor`, `server`, `cloud`, `link`, `globe`; люди — `people`, `people-list`, `user-add`, `user-growth`, `user-target`, `user-chat`, `handshake`; управление — `structure`, `list`, `layers`, `folder`, `edit`, `signature`; знаки — `check`, `bell`, `pin`, `home`, `copy`, `download`, `star`, `gear`; связь — `mail`, `send`, `phone`, `chat`.

## Портреты для титула — `"portrait": "<имя>"`

`vtb-2` — портрет с эталонной презентации «ШОиФД 2.0», используется по умолчанию на диагональном титуле.

Вырезанные портреты на прозрачном фоне, по одному на направление (`references/portraits.jpg`): `b2b`, `china`, `family`, `infra`, `private`, `privilege`, `prof`, `vtb`.

## Фото для контентных слайдов — `"photo": "<имя>"`

Обзор: `references/photos.jpg`.

| Имя | Сюжет |
|---|---|
| `lecture` | лектор у экрана, слушатели с блокнотами |
| `presenter` | двое ведущих перед аудиторией |
| `speaker` | выступающая с пультом |
| `laptop` | девушка с ноутбуком, деловой квартал |
| `team` | команда за столом (вертикальный кадр) |
| `students` | студенты общаются (вертикальный кадр) |
| `group` | студенты в учебном пространстве |
| `audience-cutout` | спикер и слушатели, вырезано (PNG) |

## Фоны-паттерны — `"background": "<направление>"`

По умолчанию берётся фон направления презентации (`references/backgrounds.jpg`): `b2b`, `china`, `family`, `infra`, `private`, `privilege`, `prof`, `vtb`.

`panel-vtb` — градиентная панель с дугами для диагонального титула и раздела (подставляется автоматически).

## Логотипы — `assets/logo/`

`logo-color.png` (светлый фон), `logo-white.png` (градиент, фото, тёмный фон), `logo-black.png` (монохром), `logo-onblue.png` (Retail Blue).

## Мокапы — `assets/mockups/`

`phone.png`, `laptop.png` — пустые экраны устройств.

## Шрифты — `assets/fonts/`

Open Sans (Light, Regular, SemiBold; лицензия OFL) — только для превью и оценки переносов. В самой презентации прописан VTB Group.
