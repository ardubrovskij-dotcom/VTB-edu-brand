"""Макеты deck.json → слайды-прототипы шаблона assets/template.pptx.

proto  — номер слайда в исходном шаблоне «мерико × ВТБ Образование» (имя слайда proto-N в template.pptx).
fields — именованные поля слайда → idx плейсхолдера.
items  — повторяющиеся элементы (карточки, пункты, хайлайты). box — область элемента в pt (x, y, w, h):
         всё, что лежит в ней, относится к элементу и удаляется, если элементов передано меньше.
         Поля элемента → idx плейсхолдера. icon: True — в области есть иконка, её можно заменить.
photo  — имя картинки-фото на прототипе (префикс), заменяется на "photo" из deck.json.
Холст 960 × 540 pt.
"""

COL4 = [(30, 100, 218, 410), (258, 100, 218, 410), (484, 100, 218, 410), (712, 100, 220, 410)]
COL3 = [(30, 100, 294, 410), (333, 100, 294, 410), (637, 100, 294, 410)]
COL2 = [(29, 100, 446, 410), (485, 100, 446, 410)]
ROWS4 = [(30, 108, 900, 90), (30, 209, 900, 90), (30, 309, 900, 90), (30, 410, 900, 90)]


def _cards(boxes, ids, photo=False):
    return [dict(box=b, title=t, text=x, icon=True, photo=photo) for b, (t, x) in zip(boxes, ids)]


LAYOUTS = {
    # ---------- обложки ----------
    'cover': dict(proto=9, fields={'title': 0, 'subtitle': 12, 'meta': 11}, portrait='Image'),
    'cover-panel': dict(proto=10, fields={'title': 0, 'subtitle': 10, 'date': 11}, portrait='Image', partner=True),
    'cover-band': dict(proto=11, fields={'title': 0, 'meta': 11}, portrait='Image'),
    'cover-text': dict(proto=12, fields={'title': 0, 'meta': 11}),
    'cover-diagonal': dict(proto=13, fields={'title': 0, 'subtitle': 10, 'date': 11}, portrait='Image', partner=True),
    'cover-photo': dict(proto=14, fields={'title': 0, 'subtitle': 10, 'date': 11}, photo='Рисунок 6', partner=True),
    'section': dict(proto=52, fields={'title': 0}, background='Рисунок 10'),

    # ---------- навигация и списки ----------
    'contents': dict(proto=15, fields={'title': 0}, items=[
        dict(box=(x, y, 293, 92), num=n, title=t, page=p) for (x, y), (n, t, p) in zip(
            [(30, 205), (334, 205), (637, 205), (30, 308), (333, 308), (637, 308), (29, 408), (333, 408), (636, 408)],
            [(30, 28, 29), (31, 32, 33), (34, 35, 36), (37, 38, 39), (40, 41, 42), (43, 44, 45),
             (46, 47, 48), (49, 50, 51), (52, 53, 54)])]),
    'bullets': dict(proto=16, fields={'title': 0}, items=[
        dict(box=b, num=n, text=t) for b, (n, t) in zip(ROWS4, [(29, 28), (31, 30), (33, 32), (35, 34)])]),
    'bullets-icons': dict(proto=17, fields={'title': 0}, items=[
        dict(box=b, text=t, icon=True) for b, t in zip(ROWS4, [28, 30, 32, 34])]),
    'bullets-conclusion': dict(proto=18, fields={'title': 0, 'conclusion': 34}, items=[
        dict(box=b, text=t, icon=True) for b, t in zip(ROWS4[:3], [28, 30, 32])]),

    # ---------- карточки ----------
    'cards2': dict(proto=19, fields={'title': 0}, items=_cards(COL2, [(55, 26), (53, 54)])),
    'cards2-photo': dict(proto=20, fields={'title': 0}, items=_cards(COL2, [(55, 26), (53, 54)], True)),
    'cards3': dict(proto=21, fields={'title': 0}, items=_cards(COL3, [(50, 26), (52, 55), (48, 54)])),
    'cards3-photo': dict(proto=22, fields={'title': 0}, items=_cards(COL3, [(50, 26), (52, 55), (48, 54)], True)),
    'cards4': dict(proto=23, fields={'title': 0}, items=_cards(COL4, [(37, 27), (39, 53), (41, 54), (43, 55)])),
    'cards4-photo': dict(proto=24, fields={'title': 0}, items=_cards(COL4, [(37, 27), (39, 53), (41, 54), (43, 55)], True)),
    'process': dict(proto=25, fields={'title': 0}, items=[
        dict(box=b, label=l, title=t, text=x) for b, (l, t, x) in zip(
            COL4, [(10, 26, 25), (33, 28, 27), (34, 30, 29), (35, 32, 31)])]),

    # ---------- карточка + фото ----------
    'photo-panel': dict(proto=26, fields={'title': 0, 'lead': 41}, photo='Рисунок 35', items=[
        dict(box=(30, y, 410, 56), text=t, icon=True) for y, t in [(255, 37), (313, 38), (370, 39), (428, 40)]]),
    'photo-list': dict(proto=27, fields={'title': 0, 'header': 21}, photo='Рисунок 14', items=[
        dict(box=(30, y, 357, 74), text=t, icon=True) for y, t in [(201, 22), (277, 24), (353, 23), (428, 25)]]),
    'photo-cards': dict(proto=28, fields={'title': 0}, photo='Рисунок 20', items=[
        dict(box=(30, y, 342, 90), title=t, text=x, icon=True) for y, (t, x) in
        [(109, (27, 26)), (210, (25, 24)), (311, (23, 22)), (411, (21, 16))]]),
    'photo-left': dict(proto=34, fields={'title': 0, 'header': 14}, photo='Рисунок 33', items=[
        dict(box=(29, y, 445, 64), text=t, icon=True) for y, t in [(170, 16), (234, 17), (299, 18), (364, 19)]]),
    'photo-right': dict(proto=35, fields={'title': 0, 'header': 14}, photo='Рисунок 26', items=[
        dict(box=(486, y, 445, 64), text=t, icon=True) for y, t in [(170, 16), (234, 17), (299, 18), (364, 19)]]),
    'laptop': dict(proto=32, fields={'title': 0, 'header': 20, 'text': 16, 'text2': 17}),

    # ---------- хайлайты (цифры) ----------
    'highlights6': dict(proto=29, fields={'title': 0}, items=[
        dict(box=(x, y, 293, 191), value=v, text=t, icon=True) for (x, y), (v, t) in zip(
            [(30, 108), (333, 108), (637, 108), (30, 309), (333, 309), (637, 308)],
            [(22, 21), (None, 25), (24, 23), (27, 26), (None, 28), (29, 30)])]),
    'highlights8': dict(proto=30, fields={'title': 0}, items=[
        dict(box=(x, y, 217, 191), value=v, text=t, icon=True) for (x, y), (v, t) in zip(
            [(30, 108), (257, 108), (484, 108), (712, 108), (29, 309), (257, 309), (484, 309), (712, 309)],
            [(22, 21), (None, 27), (24, 23), (26, 25), (29, 28), (None, 34), (31, 30), (33, 32)])]),
    'card-highlights6': dict(proto=31, fields={'title': 0, 'card_value': 22, 'card_text': 21},
                             highlight='Рисунок 49', items=[
        dict(box=(x, y, 217, 191), value=v, text=t, icon=True) for (x, y), (v, t) in zip(
            [(257, 108), (484, 108), (712, 108), (257, 309), (484, 309), (712, 309)],
            [(None, 27), (24, 23), (26, 25), (None, 34), (31, 30), (33, 32)])]),

    # ---------- таблицы ----------
    'table': dict(proto=37, fields={'title': 0, 'conclusion_title': 11, 'conclusion': 12}, table=True,
                  optional_box=(30, 423, 522, 77)),
    'table-highlight': dict(proto=38, fields={'title': 0, 'card_value': 15, 'card_text': 14}, table=True,
                            highlight='Рисунок 9'),

    # ---------- схемы и диаграммы ----------
    'scheme': dict(proto=39, fields={
        'title': 0, 'top1_title': 24, 'top1_text': 11, 'top2_title': 29, 'top2_text': 28,
        'box1_title': 16, 'box1_text': 17, 'box2_num': 32, 'box2_title': 30, 'box2_text': 31,
        'box3_num': 35, 'box3_title': 33, 'box3_text': 34, 'box4_text': 36, 'footer': 37}),
    'tree': dict(proto=40, fields={
        'title': 0, 'root_title': 10, 'root_text': 11, 'branch1_title': 24, 'branch1_text': 25,
        'leaf1': 13, 'branch2_title': 16, 'branch2_text': 17, 'leaf2': 26}),
    'venn': dict(proto=41, fields={
        'title': 0, 'left1_title': 16, 'left1_text': 17, 'left2_title': 34, 'left2_text': 35,
        'c1_title': 28, 'c1_text': 29, 'c2_title': 30, 'c2_text': 31, 'c3_title': 32, 'c3_text': 33,
        'right_title': 36, 'right_text': 37, 'card_value': 38, 'card_text': 39}),
    'charts2': dict(proto=43, fields={'title': 0, 'chart1_title': 28, 'chart2_title': 29,
                                      'card_title': 30, 'card_text': 31}, charts=['Диаграмма 45', 'Диаграмма 25']),
    'donut': dict(proto=44, fields={'title': 0, 'chart_title': 27, 'caption': 25, 'side_title': 22,
                                    'side_subtitle': 23, 'side_value_text': 28, 'side_text': 26},
                  charts=['Диаграмма 35'], textbox={'side_value': 'Текст 73'}, drop=['Текст 69', 'Полилиния']),
    'matrix': dict(proto=47, fields={'title': 0, 'x_title': 36, 'y_title': 37, 'y_top': 62, 'y_bottom': 49,
                                     'x_right': 63, 'x_title2': 64}, items=[
        dict(box=b, text=t, icon=True) for b, t in [((146, 157, 328, 166), 65), ((485, 157, 328, 166), 66),
                                                    ((146, 334, 328, 166), 67), ((485, 334, 328, 166), 68)]]),
    'timeline': dict(proto=48, fields={'title': 0}, items=[
        dict(box=b, year=y, title=t, text=x) for b, (y, t, x) in zip(
            [(80, 85, 105, 440), (186, 85, 98, 440), (285, 85, 98, 440), (384, 85, 100, 440),
             (485, 85, 118, 440), (604, 85, 108, 440), (713, 85, 170, 440)],
            [(29, 44, 31), (32, 56, 55), (34, 46, 45), (36, 54, 53), (38, 48, 47), (40, 52, 51), (42, 50, 49)])]),
    'contacts': dict(proto=49, fields={'title': 0}, card='Рисунок 524', items=[
        dict(box=(30, 108, 905, 196), name=29, role=28, phone=30, email=31, qr_text=32, qr='Группа 18', qr_box=(752, 106, 180, 198),
             qr_panel='Прямоугольник: скругленные углы 523'),
        dict(box=(30, 309, 905, 192), name=34, role=33, phone=35, email=36, qr_text=37, qr='Группа 261', qr_box=(752, 306, 180, 196))]),
}

PEOPLE = {  # ключ файла assets/people → (имя, кто на фото, направление по умолчанию)
    'alina': ('Алина', 'девушка 20–25 лет, светлые волосы, белая рубашка', 'vtb'),
    'sonya': ('Соня', 'девушка-подросток в сером худи', 'family'),
    'artyom': ('Артём', 'улыбчивый молодой человек, джинсовая рубашка', 'prof'),
    'maxim': ('Максим', 'мужчина около 30 лет, щетина, голубая рубашка', 'china'),
    'denis': ('Денис', 'мужчина с бородой, рабочая куртка', 'b2b'),
    'sergey': ('Сергей', 'мужчина 45–50 лет, белая рубашка', 'infra'),
    'viktor': ('Виктор', 'мужчина 55–60 лет с сединой, костюм и галстук', 'private'),
    'elena': ('Елена', 'женщина 40–45 лет, рыжеватые волосы, тёмный жакет', 'privilege'),
    'daria': ('Дарья', 'девушка 20–25 лет, тёмные волосы, тёмный жакет', None),
}

DIRECTIONS = {  # ключ → (номер мастера и фона/карточки, название)
    'vtb': (1, 'ВТБ Образование'),
    'family': (2, 'Финансовая грамотность для всей семьи'),
    'prof': (3, 'Профессиональные финансы'),
    'china': (4, 'Китайское направление'),
    'b2b': (5, 'B2B-направление'),
    'infra': (6, 'Крупные инфраструктурные проекты в регионах РФ'),
    'private': (7, 'Курсы по инвестициям для клиентов Private Banking'),
    'privilege': (8, 'Курсы по инвестициям для клиентов Привилегия'),
}
