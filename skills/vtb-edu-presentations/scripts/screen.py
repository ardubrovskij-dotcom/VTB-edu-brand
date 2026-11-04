"""Макеты для эфира и показа экрана (Телемост, Zoom): крупный текст, одна идея на слайд.

Кегли не меньше требований к показу экрана: заголовок 36 pt, основной текст 24 pt, запросы и подписи 20 pt.
Слайды строятся на том же прототипе, что фирменный стиль (логотип, номер), цвета — направления колоды.
Номер слайда можно задать текстом (`number`: «9а») — для кадров одного слайда сценария.

Макеты: hosts, numbered, tiles, prompt, prompt-compare, screenshot, split, statement, before-after, chat,
columns, equation, stack, list-process, questions. Поля — references/layouts.md, раздел «Эфир».
"""
import re

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

import house as H
from house import (B, CARD, E, INK, MUTED, TEXT, X0, X1, W, clear_slide, fit_size, header, items_of, lines,
                   picture_fill, rect, text, warn)

TITLE = 36      # заголовок
BODY = 24       # основной текст
SMALL = 20      # запросы, подписи, контакты — минимум для показа экрана
TOP_LIMIT = 500  # низ контента
FRAME_FILL = 'E9ECF2'
FRAME_LINE = 'ACB6C3'
# цвета частей запроса (роль, контекст, задача, формат, правила проверки) — акценты мастеров шаблона бренда
PART_COLOURS = ['1E4BD2', '00AAFF', '803CE9', 'D3875F', '0A2896']


def accent(ctx):
    return H.ACCENT[ctx['dir_no']]


def colour_of(ctx, it, k):
    c = it.get('color') if isinstance(it, dict) else None
    if isinstance(c, int):
        return PART_COLOURS[(c - 1) % len(PART_COLOURS)]
    if isinstance(c, str) and re.fullmatch(r'[0-9A-Fa-f]{6}', c):
        return c.upper()
    return PART_COLOURS[k % len(PART_COLOURS)] if c is True else None


def head(slide, s, ctx, where, max_lines=2):
    """Заголовок-вывод 36 pt (до 2 строк). Возвращает низ шапки."""
    hb = header(slide, s, ctx, where, size=TITLE, max_lines=max_lines, width=760)
    return hb


def outline(sh, colour=FRAME_LINE, width=1.5, dash=True):
    sh.line.color.rgb = RGBColor.from_string(colour)
    sh.line.width = Pt(width)
    if dash:
        sh.line.dash_style = MSO_LINE.DASH


def frame(slide, ctx, x, y, w, h, label, where, image=None, name='Рамка скриншота'):
    """Рамка-заглушка под скриншот («Скриншот: …») или сама картинка, если она уже есть."""
    path = B().asset('photos', image) if image else None
    if path:
        pic = picture_fill(slide, ctx, path, x, y, w, h, radius=6, focus=(0.5, 0.0))
        pic.name = name
        return pic
    if image:
        warn(ctx, where, f'картинка «{image}» не найдена — стоит заглушка')
    box = rect(slide, x, y, w, h, fill=FRAME_FILL)
    outline(box)
    box.name = name
    if label:
        pad = 16 if w < 300 else 30
        # в узкой рамке подпись мельче: длинное слово («обезличивания») не должно рваться
        ls, ok = fit_size(ctx, label, BODY if w >= 300 else SMALL, (w - 2 * pad) * 0.92, h - 20, floor=12)
        text(slide, x + pad, y, w - 2 * pad, h, label, ls, MUTED, anchor='middle', align='center', pitch=ls * 1.2,
             name='Подпись рамки')
    return box


def badge(slide, ctx, x, y, d, label, fill, fg='FFFFFF', size=SMALL):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int(x * E)), Emu(int(y * E)), Emu(int(d * E)), Emu(int(d * E)))
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    sh.line.fill.background()
    sh.shadow.inherit = False
    sh.name = 'Номер'
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(label)
    r.font.size = Pt(size)
    r.font.name = H.FONT_DEMI
    r.font.color.rgb = RGBColor.from_string(fg)
    return sh


def centred(top, need, bottom=TOP_LIMIT):
    """Верх блока высотой need, поставленного по центру свободного поля под шапкой."""
    return max(top, top + (bottom - top - need) / 2)


def each_fit(ctx, where, paras, width, height, size=BODY, what='текст'):
    """Один кегль для нескольких отдельных блоков (карточек, столбцов): самый мелкий из тех, что нужны каждому."""
    s = size
    for p in paras:
        if p:
            s = min(s, body_fit(ctx, where, p, width, height, size, what=what, report=False))
    if any(p and not fit_size(ctx, p, s, width, height, 1.2, floor=s)[1] for p in paras):
        warn(ctx, where, f'ПЕРЕПОЛНЕНИЕ — {what} не помещается даже при {SMALL} pt: сократить или разбить на кадры')
    return s


def body_fit(ctx, where, paras, width, height, size=BODY, gap=0, what='текст', report=True):
    s, ok = fit_size(ctx, paras, size, width, height, 1.2, floor=SMALL if size >= SMALL else size, gap=gap)
    if not ok and report:
        warn(ctx, where, f'ПЕРЕПОЛНЕНИЕ — {what} не помещается даже при {SMALL} pt: сократить или разбить на кадры')
    return s


# ---------------------------------------------------------------- макеты
def h_hosts(slide, s, ctx, where):
    """Приветствие: название эфира, дата, крупная надпись «Роли» и карточки ведущих (фото, имя, о себе, роль, почта)."""
    clear_slide(slide)
    t = next((p for p in slide.placeholders if p.placeholder_format.type == 1), None)
    if t is not None:
        t._element.getparent().remove(t._element)
    col = accent(ctx)
    title = s.get('title', '')
    ts, _ = fit_size(ctx, title, 44, 740, 110, 1.1, floor=TITLE)
    tn = lines(ctx, title, ts, 740)
    text(slide, X0, 30, 740, tn * ts * 1.1, title, ts, col, pitch=ts * 1.1, name='Название')
    y = 30 + tn * ts * 1.1 + 6
    if s.get('date'):
        text(slide, X0, y, 740, 30, s['date'], BODY, MUTED, name='Дата')
        y += 34
    its = items_of(s)[:2]
    roles = s.get('roles_title', 'Роли')
    text(slide, X0, y + 14, 400, 40, roles, 32, col, bold=True, name='Роли')
    top = y + 14 + 46
    n = max(1, len(its))
    w = (W - 10 * (n - 1)) / n
    h = TOP_LIMIT - top
    for k, it in enumerate(its):
        x = X0 + k * (w + 10)
        rect(slide, x, top, w, h)
        ph = min(130, h * 0.47)
        pw = ph * 0.9
        frame(slide, ctx, x + 18, top + 18, pw, ph, 'фото', where, it.get('photo'), name='Фото ведущего')
        tx = x + 18 + pw + 16
        tw = x + w - 18 - tx
        nm = it.get('name', '[имя и фамилия]')
        nl = lines(ctx, nm, 28, tw, bold=True)  # имя — 28 pt, как на финальном слайде; можно в две строки
        text(slide, tx, top + 18, tw, nl * 28 * 1.15, nm, 28, INK, bold=True, pitch=28 * 1.15, name='Имя')
        ay = top + 18 + nl * 28 * 1.15 + 6
        if it.get('about'):
            ab = body_fit(ctx, where, it['about'], tw, top + 18 + ph - ay, SMALL, what='строка о себе')
            text(slide, tx, ay, tw, top + 18 + ph - ay, it['about'], ab, TEXT, pitch=ab * 1.2, name='О себе')
        ry = top + 18 + ph + 16
        if it.get('role'):
            rs = body_fit(ctx, where, it['role'], w - 36, h - (ry - top) - 40, BODY, what='роль')
            rn = lines(ctx, it['role'], rs, w - 36)
            text(slide, x + 18, ry, w - 36, rn * rs * 1.2, it['role'], rs, col, pitch=rs * 1.2, name='Роль')
        if it.get('email'):
            text(slide, x + 18, top + h - 18 - 26, w - 36, 26, it['email'], SMALL, TEXT, name='Почта')


def h_numbered(slide, s, ctx, where):
    """Нумерованный список 24 pt: номера в кружках; group — подпись группы, accent — выделенный пункт;
    columns: 2 — в две колонки."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    its = items_of(s)
    if len(its) > 5 and not s.get('columns'):
        warn(ctx, where, f'в списке {len(its)} пунктов — больше пяти, проверить по ТЗ')
    cols = s.get('columns') or 1
    per = (len(its) + cols - 1) // cols
    colw = (W - 30 * (cols - 1)) / cols
    d = 38
    tw = colw - d - 18 - (24 if any(it.get('group') for it in its) else 0)
    texts = [it.get('text', '') for it in its]
    size = body_fit(ctx, where, texts, tw, (TOP_LIMIT - hb - 30) / max(1, per) * 0.8 * per, BODY, what='список')
    col = accent(ctx)
    rows = []
    for it in its:
        n = lines(ctx, it.get('text', ''), size, tw)
        rows.append(max(d, n * size * 1.2) + (30 if it.get('group') else 0))
    gap = 14
    for c in range(cols):
        chunk = list(range(c * per, min(len(its), (c + 1) * per)))
        need = sum(rows[i] for i in chunk) + gap * (len(chunk) - 1)
        y = centred(hb + 24, need)
        x = X0 + c * (colw + 30)
        group_start = None
        for i in chunk:
            it = its[i]
            num = it.get('num', i + 1)
            gx = x + (24 if any(j.get('group') for j in its) else 0)
            if it.get('group'):
                text(slide, x, y, colw, 26, it['group'], SMALL, MUTED, bold=True, name='Подпись группы')
                if group_start is not None:
                    bar(slide, x + 6, group_start, y - 10 - group_start, col)
                group_start = y + 30
                y += 30
            h = rows[i] - (30 if it.get('group') else 0)
            if it.get('accent'):
                hl = rect(slide, gx - 10, y - 8, colw - (gx - x) + 10, h + 16, fill='EAF0FF')
                outline(hl, col, 2, dash=False)
                hl.name = 'Выделение'
            badge(slide, ctx, gx, y + (min(h, d) - d) / 2, d, num, col)
            text(slide, gx + d + 18, y + (d - size * 1.2) / 2 if h <= d else y, tw, h, it.get('text', ''), size,
                 INK, pitch=size * 1.2, name='Пункт')
            y += h + gap
        if group_start is not None:
            bar(slide, x + 6, group_start, y - gap - group_start, col)


def bar(slide, x, y, h, colour):
    if h > 4:
        b = rect(slide, x, y, 4, h, fill=colour, radius=0)
        b.name = 'Скобка группы'


def h_tiles(slide, s, ctx, where):
    """Карточки 2–6 (по 3 в ряд): цветная полоса слева (`color`: номер части 1–5 или HEX), заголовок 28 pt
    цветом карточки, текст 24 pt."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    its = items_of(s)[:6]
    n = len(its)
    cols = s.get('columns') or (n if n <= 3 else 3 if n in (5, 6) else 2)
    rows = (n + cols - 1) // cols
    w = (W - 10 * (cols - 1)) / cols
    top = hb + 28
    h = min(200, (TOP_LIMIT - top - 10 * (rows - 1)) / rows)
    top = centred(top, rows * h + 10 * (rows - 1))
    pad = 22
    bodies = [it.get('text', '') for it in its]
    tw = w - 2 * pad - 8
    tl = max([lines(ctx, it.get('title', ''), 28, tw, bold=True) for it in its] or [1])  # строк в заголовках
    th = tl * 28 * 1.15
    ts = each_fit(ctx, where, bodies, tw, h - 2 * pad - th - 8, BODY, what='текст карточки')
    for k, it in enumerate(its):
        x = X0 + (k % cols) * (w + 10)
        y = top + (k // cols) * (h + 10)
        c = colour_of(ctx, it, k) or accent(ctx)
        rect(slide, x, y, w, h)
        stripe = rect(slide, x, y, 8, h, fill=c, radius=0)
        stripe.name = 'Цвет части'
        text(slide, x + pad + 8, y + pad, tw, th, it.get('title', ''), 28, c, bold=True, pitch=28 * 1.15,
             name='Заголовок карточки')
        if it.get('text'):
            text(slide, x + pad + 8, y + pad + th + 8, tw, h - 2 * pad - th - 8, it['text'], ts, INK,
                 pitch=ts * 1.2, name='Текст карточки')


def code_box(slide, x, y, w, h, name='Запрос'):
    box = rect(slide, x, y, w, h, fill='F7F8FB')
    outline(box, 'C9CFDA', 1.25, dash=False)
    box.name = name
    return box


def h_prompt(slide, s, ctx, where):
    """Запрос в рамке блока кода, 20 pt. parts — части запроса с подписями и цветами (как на карточках `tiles`)."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    top = hb + 22
    h = TOP_LIMIT - top
    code_box(slide, X0, top, W, h)
    parts = s.get('parts')
    pad = 22
    if not parts:
        body = s.get('prompt', '[запрос — дословно из сценария]')
        ps = body_fit(ctx, where, body, W - 2 * pad, h - 2 * pad, SMALL, what='запрос')
        text(slide, X0 + pad, top + pad, W - 2 * pad, h - 2 * pad, body, ps, INK, pitch=ps * 1.25, name='Запрос')
        return
    lw = 170
    tw = W - 2 * pad - lw - 16
    bodies = [p.get('text', '') for p in parts]
    gap = 12
    ps = body_fit(ctx, where, bodies, tw, h - 2 * pad - gap * (len(parts) - 1), SMALL, gap=gap, what='запрос')
    y = top + pad
    for k, p in enumerate(parts):
        c = colour_of(ctx, {'color': p.get('color', k + 1)}, k)
        n = lines(ctx, p.get('text', ''), ps, tw)
        ph = max(n * ps * 1.25, 26)
        stripe = rect(slide, X0 + pad, y, 6, ph, fill=c, radius=0)
        stripe.name = 'Цвет части'
        text(slide, X0 + pad + 14, y, lw - 14, ph, p.get('label', ''), SMALL, c, bold=True, name='Часть запроса')
        text(slide, X0 + pad + lw + 16, y, tw, ph, p.get('text', ''), ps, INK, pitch=ps * 1.25, name='Запрос')
        y += ph + gap


def h_prompt_compare(slide, s, ctx, where):
    """Два запроса рядом (слабый слева, сильный справа), под ними вопросы для сравнения."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    its = items_of(s)[:2]
    qs = s.get('questions') or []
    qh = 0
    if qs:
        qh = 92
    top = hb + 22
    capt = 34
    h = TOP_LIMIT - top - capt - (qh + 14 if qs else 0)
    w = (W - 20) / 2
    bodies = [it.get('text', '') for it in its]
    ps = each_fit(ctx, where, bodies, w - 40, h - 36, SMALL, what='запрос')
    for k, it in enumerate(its):
        x = X0 + k * (w + 20)
        text(slide, x, top, w, capt, it.get('title', ''), BODY, accent(ctx) if k else MUTED, bold=True,
             name='Подпись запроса')
        code_box(slide, x, top + capt, w, h)
        text(slide, x + 20, top + capt + 18, w - 40, h - 36, it.get('text', ''), ps, INK, pitch=ps * 1.25,
             name='Запрос')
    if qs:
        y = TOP_LIMIT - qh
        qw = (W - 20) / len(qs)
        qsz = each_fit(ctx, where, qs, qw - 72, qh - 16, SMALL, what='вопросы')
        for k, q in enumerate(qs):
            x = X0 + k * (qw + 10)
            rect(slide, x, y, qw, qh, fill='EAF0FF')
            badge(slide, ctx, x + 14, y + (qh - 34) / 2, 34, '?', accent(ctx))
            text(slide, x + 60, y, qw - 72, qh, q, qsz, INK, anchor='middle', pitch=qsz * 1.2, name='Вопрос')


def h_screenshot(slide, s, ctx, where):
    """Рамка под скриншот 16 : 9 (не меньше 60 % площади слайда) с подписью «Скриншот: …»; под рамкой — note
    («результат из заготовки»). image — готовый снимок вместо заглушки."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where, max_lines=1)
    note = s.get('note')
    y = hb + 10
    h = TOP_LIMIT - y               # 60 % площади слайда; подпись под рамкой — на линии номера слайда
    w = min(W, h * 16 / 9)
    h = w * 9 / 16
    x = X0 + (W - w) / 2
    area = w * h / (960 * 540)
    if area < 0.6:
        warn(ctx, where, f'рамка скриншота {area:.0%} площади слайда — меньше 60 % по ТЗ; сократить заголовок до строки')
    frame(slide, ctx, x, y, w, h, s.get('label', 'Скриншот: [что именно]'), where, s.get('image'))
    if note:
        text(slide, x, y + h + 4, min(w, 820 - x), 26, note, SMALL, MUTED, name='Подпись под рамкой')


def h_split(slide, s, ctx, where):
    """Слева список (нумерованный или поля шаблона «название: значение»), справа рамка под скриншот или
    изображение. side: right/left — где рамка."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    top = hb + 22
    h_av = TOP_LIMIT - top - (30 if s.get('note') else 0)
    fw = s.get('frame_width') or 560
    fh = min(h_av, fw * 9 / 16) if s.get('ratio', '16:9') == '16:9' else h_av
    fw = fh * 16 / 9 if s.get('ratio', '16:9') == '16:9' else fw
    lw = W - fw - 30
    right = s.get('side', 'right') == 'right'
    fx = X0 + lw + 30 if right else X0
    lx = X0 if right else X0 + fw + 30
    fy = centred(top, fh, TOP_LIMIT - (30 if s.get('note') else 0))
    frame(slide, ctx, fx, fy, fw, fh, s.get('label', 'Скриншот: [что именно]'), where, s.get('image'))
    if s.get('note'):
        text(slide, fx, fy + fh + 6, fw, 26, s['note'], SMALL, MUTED, name='Подпись под рамкой')
    its = items_of(s)
    fields = any(it.get('label') for it in its)
    numbered = s.get('numbered', not fields)
    d = 34
    tw = lw - (d + 14 if numbered else 0)
    paras = [(f"{it['label']}: " if it.get('label') else '') + it.get('text', '') for it in its]
    gap = 12
    size = body_fit(ctx, where, paras, tw, h_av - gap * (len(its) - 1), BODY if len(its) <= 5 else SMALL,
                    gap=gap, what='список')
    rows = [max(d if numbered else 0, lines(ctx, p, size, tw) * size * 1.2) for p in paras]
    y = centred(top, sum(rows) + gap * (len(rows) - 1))
    col = accent(ctx)
    for k, (it, p, rh) in enumerate(zip(its, paras, rows)):
        if numbered:
            badge(slide, ctx, lx, y + (min(rh, d) - d) / 2, d, it.get('num', k + 1), col)
        tx = lx + (d + 14 if numbered else 0)
        tb = text(slide, tx, y, tw, rh, (f"**{it['label']}:** " if it.get('label') else '') + it.get('text', ''),
                  size, INK, pitch=size * 1.2, name='Пункт')
        if it.get('label'):
            tb.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string(col)
        y += rh + gap


def h_statement(slide, s, ctx, where):
    """Одна фраза на весь слайд (вопрос, вывод): крупно, по центру, без заголовка и картинок."""
    clear_slide(slide)
    t = next((p for p in slide.placeholders if p.placeholder_format.type == 1), None)
    if t is not None:
        t._element.getparent().remove(t._element)
    body = s.get('text', s.get('title', ''))
    size, ok = fit_size(ctx, body, s.get('size', 48), 800, 330, 1.15, floor=TITLE)
    if not ok:
        warn(ctx, where, 'ПЕРЕПОЛНЕНИЕ — фраза не помещается крупно, сократить')
    n = lines(ctx, body, size, 800)
    h = n * size * 1.15
    y = (540 - h) / 2 - 10
    text(slide, 80, y, 800, h + 10, body, size, accent(ctx), pitch=size * 1.15, align='center', name='Фраза')


def h_before_after(slide, s, ctx, where):
    """Было → стало: две страницы (A4, книжная) рядом с подписями и стрелкой между ними, подпись внизу."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where, max_lines=1)
    its = items_of(s)[:2]
    note = s.get('note')
    top = hb + 16
    capt = 34
    bottom = TOP_LIMIT + 6 - (30 if note else 0)
    h = bottom - top - capt
    w = h / 1.414
    gapx = 120
    x0 = (960 - (2 * w + gapx)) / 2
    for k, it in enumerate(its):
        x = x0 + k * (w + gapx)
        text(slide, x, top, w, capt, it.get('title', 'Было' if k == 0 else 'Стало'), 28, accent(ctx), bold=True,
             align='center', name='Подпись листа')
        frame(slide, ctx, x, top + capt, w, h, it.get('label', '[лист]'), where, it.get('image'))
    ar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Emu(int((x0 + w + 30) * E)),
                                Emu(int((top + capt + h / 2 - 22) * E)), Emu(int((gapx - 60) * E)), Emu(int(44 * E)))
    ar.fill.solid()
    ar.fill.fore_color.rgb = RGBColor.from_string(accent(ctx))
    ar.line.fill.background()
    ar.name = 'Стрелка'
    if note:
        text(slide, X0, bottom + 4, W, 26, note, SMALL, MUTED, align='center', name='Подпись внизу')


def h_chat(slide, s, ctx, where):
    """Слева окно чата со строками (заметки, переписка), справа пустое место под живой показ."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    top = hb + 22
    h = TOP_LIMIT - top
    w = (W - 20) / 2
    win = rect(slide, X0, top, w, h, fill='FFFFFF')
    outline(win, 'C9CFDA', 1.25, dash=False)
    win.name = 'Окно чата'
    bar_ = rect(slide, X0, top, w, 30, fill='EEF0F4')
    bar_.name = 'Шапка окна'
    for k, c in enumerate(('E3E6EC', 'E3E6EC', 'E3E6EC')):
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int((X0 + 14 + k * 16) * E)), Emu(int((top + 10) * E)),
                                     Emu(int(10 * E)), Emu(int(10 * E)))
        dot.fill.solid()
        dot.fill.fore_color.rgb = RGBColor.from_string('B4BAC6')
        dot.line.fill.background()
    msgs = s.get('lines') or []
    bw = w - 40
    gap = 8
    ms = SMALL
    for cand in (BODY, 22, SMALL):  # самый крупный кегль, при котором все строки помещаются в окно
        need = sum(lines(ctx, m, cand, bw - 24) * cand * 1.2 + 14 for m in msgs) + gap * (len(msgs) - 1)
        if need <= h - 30 - 28:
            ms = cand
            break
    y = top + 30 + 14
    for m in msgs:
        n = lines(ctx, m, ms, bw - 24)
        bh = n * ms * 1.2 + 14
        bub = rect(slide, X0 + 20, y, bw, bh, fill='EEF2FB', radius=10)
        bub.name = 'Сообщение'
        text(slide, X0 + 32, y + 7, bw - 24, bh - 14, m, ms, INK, pitch=ms * 1.2, name='Строка чата')
        y += bh + gap
    if y > top + h:
        warn(ctx, where, 'ПЕРЕПОЛНЕНИЕ — строки не помещаются в окно чата')
    rx = X0 + w + 20
    empty = rect(slide, rx, top, w, h, fill='FFFFFF')
    outline(empty)
    empty.name = 'Место для живого показа'
    text(slide, rx + 30, top, w - 60, h, s.get('right_label', 'Здесь появится результат — показываем живьём'),
         BODY, MUTED, anchor='middle', align='center', pitch=BODY * 1.2, name='Подпись места')


def h_columns(slide, s, ctx, where):
    """Два столбца-сравнения: заголовок столбца 32 pt и текст 24 pt."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    its = items_of(s)[:3]
    n = max(1, len(its))
    w = (W - 20 * (n - 1)) / n
    top = hb + 28
    h = min(300, TOP_LIMIT - top)
    top = centred(top, h)
    bodies = [it.get('text', '') for it in its]
    ts = each_fit(ctx, where, bodies, w - 60, h - 110, BODY, what='текст столбца')
    for k, it in enumerate(its):
        x = X0 + k * (w + 20)
        c = colour_of(ctx, it, k) or accent(ctx)
        rect(slide, x, top, w, h)
        text(slide, x + 30, top + 30, w - 60, 44, it.get('title', ''), 32, c, bold=True, name='Заголовок столбца')
        text(slide, x + 30, top + 90, w - 60, h - 110, it.get('text', ''), ts, INK, pitch=ts * 1.25,
             name='Текст столбца')


def h_equation(slide, s, ctx, where):
    """Схема «A + B = C»: входы и результат в карточках, между ними знаки."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    its = items_of(s)[:4]
    ops = s.get('operators') or (['+'] * (len(its) - 2) + ['='])
    n = len(its)
    opw = 60
    w = (W - opw * (n - 1)) / n
    h = 170
    top = centred(hb + 28, h)
    bodies = [it.get('text', '') for it in its]
    ts = each_fit(ctx, where, bodies, w - 40, h - 40, 28, what='текст схемы')
    for k, it in enumerate(its):
        x = X0 + k * (w + opw)
        last = k == n - 1
        if last:
            rect(slide, x, top, w, h, grad=H.GRAD[ctx['dir_no']])
        else:
            rect(slide, x, top, w, h)
        text(slide, x + 20, top, w - 40, h, it.get('text', ''), ts, 'FFFFFF' if last else INK, anchor='middle',
             align='center', bold=last, pitch=ts * 1.2, name='Элемент схемы')
        if k < n - 1:
            text(slide, x + w, top, opw, h, ops[k], 48, accent(ctx), anchor='middle', align='center', name='Знак')


def h_stack(slide, s, ctx, where):
    """Схема страницы: блоки сверху вниз в рамке страницы; последний — кнопка."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    its = items_of(s)
    top = hb + 22
    h = TOP_LIMIT - top
    pw = 560
    px = X0 + (W - pw) / 2
    page = rect(slide, px, top, pw, h, fill='FFFFFF')
    outline(page, 'C9CFDA', 1.25, dash=False)
    page.name = 'Страница'
    pad, gap = 18, 8
    bh = (h - 2 * pad - gap * (len(its) - 1)) / max(1, len(its))
    bodies = [it.get('text', '') for it in its]
    ts = each_fit(ctx, where, bodies, pw - 2 * pad - 70, bh - 6, BODY, what='блок страницы')
    col = accent(ctx)
    for k, it in enumerate(its):
        y = top + pad + k * (bh + gap)
        button = it.get('button') or (k == len(its) - 1 and s.get('last_button', True))
        if button:
            bw = 260
            btn = rect(slide, px + (pw - bw) / 2, y, bw, bh, fill=col, radius=bh / 2)
            btn.name = 'Кнопка'
            text(slide, px + (pw - bw) / 2, y, bw, bh, it.get('text', ''), ts, 'FFFFFF', anchor='middle',
                 align='center', name='Блок страницы')
        else:
            rect(slide, px + pad, y, pw - 2 * pad, bh, fill=CARD)
            badge(slide, ctx, px + pad + 12, y + (bh - 34) / 2, 34, k + 1, col)
            text(slide, px + pad + 58, y, pw - 2 * pad - 70, bh, it.get('text', ''), ts, INK, anchor='middle',
                 name='Блок страницы')


def h_list_process(slide, s, ctx, where):
    """Сверху список (до 6 пунктов в две колонки), снизу процесс из шагов со стрелками и строка-примечание."""
    clear_slide(slide)
    hb = head(slide, s, ctx, where)
    its = items_of(s)
    steps = s.get('steps') or []
    note = s.get('note')
    col = accent(ctx)
    cols = 2 if len(its) > 3 else 1
    per = (len(its) + cols - 1) // cols
    colw = (W - 30) / cols
    rowh = BODY * 1.2 + 8
    y0 = hb + 24
    for k, it in enumerate(its):
        x = X0 + (k // per) * (colw + 30)
        y = y0 + (k % per) * rowh
        dot = rect(slide, x, y + BODY * 0.45, 10, 10, fill=col, radius=5)
        dot.name = 'Маркер'
        text(slide, x + 22, y, colw - 22, rowh, it.get('text', ''), BODY, INK, name='Пункт')
    y = y0 + per * rowh + 18
    if s.get('steps_title'):
        text(slide, X0, y, W, 32, s['steps_title'], BODY, col, bold=True, name='Подпись процесса')
        y += 40
    if steps:
        n = len(steps)
        aw = 44
        sw = (W - aw * (n - 1)) / n
        sh_ = 96
        ss = each_fit(ctx, where, steps, sw - 40, sh_ - 20, BODY, what='шаги')
        for k, st in enumerate(steps):
            x = X0 + k * (sw + aw)
            rect(slide, x, y, sw, sh_, fill='EAF0FF')
            text(slide, x + 22, y, sw - 44, sh_, st, ss, INK, anchor='middle', align='center', pitch=ss * 1.2,
                 name='Шаг')
            if k < n - 1:
                ar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Emu(int((x + sw + 8) * E)),
                                            Emu(int((y + sh_ / 2 - 12) * E)), Emu(int((aw - 16) * E)),
                                            Emu(int(24 * E)))
                ar.fill.solid()
                ar.fill.fore_color.rgb = RGBColor.from_string(col)
                ar.line.fill.background()
        y += sh_ + 14
    if note:
        text(slide, X0, y, W, 28, note, SMALL, MUTED, name='Примечание')
    if y + 28 > TOP_LIMIT + 10:
        warn(ctx, where, 'ПЕРЕПОЛНЕНИЕ — список и процесс не помещаются')


def h_questions(slide, s, ctx, where):
    """Финал: крупное слово, контакты ведущих, рамка под QR-код."""
    clear_slide(slide)
    t = next((p for p in slide.placeholders if p.placeholder_format.type == 1), None)
    if t is not None:
        t._element.getparent().remove(t._element)
    col = accent(ctx)
    text(slide, X0, 110, 560, 90, s.get('title', 'Вопросы'), 72, col, name='Название')
    y = 230
    for it in items_of(s)[:3]:
        text(slide, X0, y, 560, 32, it.get('name', ''), 28, INK, bold=True, name='Имя')
        if it.get('email'):
            text(slide, X0, y + 36, 560, 28, it['email'], BODY, TEXT, name='Почта')
        y += 84
    q = 230
    qx, qy = X1 - q, 150
    frame(slide, ctx, qx, qy, q, q, s.get('qr_label', 'QR-код'), where, s.get('qr'), name='QR-код')
    if s.get('qr_caption'):
        text(slide, qx - 40, qy + q + 10, q + 40, 60, s['qr_caption'], SMALL, MUTED, align='center',
             pitch=SMALL * 1.2, name='Подпись QR')


SCREEN = {
    'hosts': h_hosts, 'numbered': h_numbered, 'tiles': h_tiles, 'prompt': h_prompt,
    'prompt-compare': h_prompt_compare, 'screenshot': h_screenshot, 'split': h_split, 'statement': h_statement,
    'before-after': h_before_after, 'chat': h_chat, 'columns': h_columns, 'equation': h_equation,
    'stack': h_stack, 'list-process': h_list_process, 'questions': h_questions,
}
