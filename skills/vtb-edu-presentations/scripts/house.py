"""Фирменный стиль команды — макеты по эталону «ШОиФД 2.0» (references/etalon-shoifd.pptx, финальная версия автора).

Слайд строится на прототипе шаблона (логотип, номер слайда, мастер направления), а содержимое рисуется заново
по размерам эталона. Холст 960 × 540 pt.

Общие правила эталона:
- заголовок — вывод слайда целиком, обычным регистром, цветом направления, 28 pt, шаг 31 pt, до 3 строк;
- подзаголовок — серым 19 pt сразу под заголовком; сноска — 9,5 pt внизу слева;
- контент прижат к низу: нижняя граница блоков y = 487, между шапкой и контентом — свободное поле;
- карточки серые #F4F6FA, скругление 6 pt, шаг 10 pt, без шапок; акцентная карточка — градиент направления;
- цифры — 42 pt цветом направления, слова и «/5» рядом мелко («17 стран», «80 из 89», «4,7/5»).
"""
import copy
import os
import re

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.util import Emu, Pt

E = 12700
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

X0, X1, W = 30, 930, 900          # поля и ширина контента
BOTTOM = 487                      # нижняя граница контента
GAP = 10                          # шаг между карточками
TITLE_TOP, TITLE_SIZE, TITLE_PITCH = 30, 28, 31  # заголовок до 3 строк (правка автора)
SUB_SIZE, SUB_PITCH = 19, 23  # подзаголовок крупнее (правка автора)
CARD = 'F4F6FA'
TEXT = '3A4250'
MUTED = '6B7683'
NOTE = '4A5361'                   # сноски, подпись «Продолжительность»
VALUE_SIZE = 42                   # цифры в плитках и над спикерами
SMALL_IN_VALUE = r'[А-Яа-яЁё]+|/\s*\d+'  # мелко внутри цифры: «из», «стран», «/5», «/10»
INK = '1C1C1C'
FONT = 'VTB Omega UI'
FONT_DEMI = 'VTB Omega UI Medium'  # выделение — Medium: DemiBold и Bold слишком жирные (правка автора)

# цвет цифр и имён, градиент акцентных блоков (слева направо) — по направлению
ACCENT = {1: '0A2896', 2: '803CE9', 3: '750CF6', 4: '0A2896', 5: '1E4BD2', 6: '1E4BD2', 7: '000324', 8: '900258'}
GRAD = {1: ('1E4BD2', '86BFF5'), 2: ('00AAFF', '803CE9'), 3: ('0A2896', '750CF6'), 4: ('0A2896', 'D8A2F7'),
        5: ('1E4BD2', 'FFB68E'), 6: ('1E4BD2', '00AAFF'), 7: ('000324', '57637A'), 8: ('002043', '900258')}


def B():
    import build_deck
    return build_deck


# ---------------------------------------------------------------- примитивы
def rect(slide, x, y, w, h, fill=CARD, radius=6, grad=None, angle=0):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(int(x * E)), Emu(int(y * E)),
                                Emu(int(w * E)), Emu(int(h * E)))
    sh.adjustments[0] = min(0.5, radius / max(1, min(w, h)))
    sh.line.fill.background()
    sh.shadow.inherit = False
    if grad:
        sh.fill.gradient()
        sh.fill.gradient_angle = angle
        stops = sh.fill.gradient_stops
        stops[0].color.rgb = RGBColor.from_string(grad[0])
        stops[0].position = 0
        stops[1].color.rgb = RGBColor.from_string(grad[1])
        stops[1].position = 1.0
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    sh.text_frame.text = ''
    sh.name = 'Карточка'
    return sh


def runs_of(text):
    """'**жирное** обычное' → [(текст, жирный)]."""
    out = []
    for i, part in enumerate(re.split(r'\*\*', text)):
        if part:
            out.append((part, i % 2 == 1))
    return out or [('', False)]


def text(slide, x, y, w, h, paras, size=14, color=TEXT, bold=False, anchor='top', pitch=None, caps=False,
         align='left', bullet=None, gap=0, small=None, name='Текст', numbered=False):
    """Текстовое поле без отступов. paras — строка или список абзацев. bullet — цвет маркера «•».
    small — (кегль, regex): совпадения набираются мелким кеглем (единицы у цифр)."""
    tb = slide.shapes.add_textbox(Emu(int(x * E)), Emu(int(y * E)), Emu(int(w * E)), Emu(int(h * E)))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {'top': MSO_ANCHOR.TOP, 'bottom': MSO_ANCHOR.BOTTOM, 'middle': MSO_ANCHOR.MIDDLE}[anchor]
    if isinstance(paras, str):
        paras = paras.replace('\\n', '\n').split('\n')
    nbsp = B().nbsp
    for i, ptxt in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}[align]
        p.line_spacing = Pt(pitch or size * 1.2)
        if gap and i:
            p.space_before = Pt(gap)
        if numbered:  # автонумерация «1.» с висячим отступом
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', str(int(16 * E)))
            pPr.set('indent', str(int(-16 * E)))
            pPr.append(parse_xml(f'<a:buFont xmlns:a="{A}" typeface="+mj-lt"/>'))
            pPr.append(parse_xml(f'<a:buAutoNum xmlns:a="{A}" type="arabicPeriod"/>'))
        elif bullet:
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', str(int(14 * E)))
            pPr.set('indent', str(int(-14 * E)))
            pPr.append(parse_xml(f'<a:buClr xmlns:a="{A}"><a:srgbClr val="{bullet}"/></a:buClr>'))
            pPr.append(parse_xml(f'<a:buFont xmlns:a="{A}" typeface="Arial"/>'))
            pPr.append(parse_xml(f'<a:buChar xmlns:a="{A}" char="•"/>'))
        pieces = []
        for seg, b in runs_of(nbsp(ptxt.upper() if caps else ptxt)):
            if small:
                last = 0
                for m in re.finditer(small[1], seg):
                    if m.start() > last:
                        pieces.append((seg[last:m.start()], b, size))
                    pieces.append((m.group(0), b, small[0]))
                    last = m.end()
                if last < len(seg):
                    pieces.append((seg[last:], b, size))
            else:
                pieces.append((seg, b, size))
        for seg, b, sz in pieces:
            r = p.add_run()
            r.text = seg
            f = r.font
            f.size = Pt(sz)
            f.name = FONT_DEMI if (bold or b) else FONT
            f.color.rgb = RGBColor.from_string(color)
    return tb


def lines(ctx, paras, size, width, caps=False, bold=False):
    if isinstance(paras, str):
        paras = paras.replace('\\n', '\n').split('\n')
    return sum(ctx['fitter'].lines(re.sub(r'\*\*', '', p), size, width, caps, bold or '**' in p) for p in paras)


def fit_size(ctx, paras, size, width, height, pitch_k=1.2, floor=9, caps=False, gap=0, role=None):
    """Кегль, при котором абзацы помещаются в рамку; второе значение — поместились ли.
    role — элемент для единого кегля по колоде (см. UNIFORM в build_deck.py)."""
    n_par = len(paras) if isinstance(paras, list) else 1
    s = B().capped(role, size)
    words = re.findall(r'\S+', re.sub(r'\*\*', '', ' '.join(paras) if isinstance(paras, list) else paras))
    while True:
        h = lines(ctx, paras, s, width, caps) * s * pitch_k + gap * (n_par - 1)
        # самое длинное слово тоже должно помещаться, иначе оно разорвётся по буквам
        widest = max((ctx['fitter'].width(w, s) for w in words), default=0)
        if (h <= height and widest <= width) or s <= floor:
            B().note_size(role, size, s)
            return s, h <= height
        s -= 0.5


def warn(ctx, where, msg):
    ctx['report'].append(f'{where}: {msg}')


def clear_slide(slide):
    """Убирает с прототипа всё, кроме заголовка и номера слайда."""
    for sh in list(slide.shapes):
        if sh.is_placeholder and sh.placeholder_format.type in (1, 13):
            continue
        sh._element.getparent().remove(sh._element)


def header(slide, s, ctx, where, top=TITLE_TOP, width=735, size=TITLE_SIZE, max_lines=3):
    """Заголовок-вывод и подзаголовок. Возвращает нижнюю границу шапки."""
    title = s.get('title', '')
    t = next((p for p in slide.placeholders if p.placeholder_format.type == 1), None)
    role = 'заголовок' if size == TITLE_SIZE else None
    sz = B().capped(role, size)
    n = lines(ctx, title, sz, width, caps=False) if title else 0
    while n > max_lines and sz > 20:
        sz -= 1
        n = lines(ctx, title, sz, width, caps=False)
    B().note_size(role, size, sz)
    if n > max_lines:
        warn(ctx, where, f'ПЕРЕПОЛНЕНИЕ — заголовок {n} строк при пределе {max_lines}. Сократите его')
    if sz < size:
        warn(ctx, where, f'кегль заголовка {size} → {sz} pt')
    pitch = sz * TITLE_PITCH / TITLE_SIZE
    y = top + n * pitch
    if t is not None:
        if title:
            t.left, t.top, t.width, t.height = (Emu(int(v * E)) for v in (X0, top, width, n * pitch + 6))
            tf = t.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.TOP
            no_autoshrink(tf)
            tf.text = B().nbsp(title)
            for p in tf.paragraphs:
                p.line_spacing = Pt(pitch)
                for r in p.runs:
                    r.font.size = Pt(sz)
                    r.font.color.rgb = RGBColor.from_string(ACCENT[ctx['dir_no']])  # заголовки — цветом направления
        else:
            t._element.getparent().remove(t._element)
    if s.get('title2'):  # вторая строка заголовка капсом, мельче («Базовый формат:» + суть)
        n2 = lines(ctx, s['title2'], 16, width, caps=False)
        text(slide, X0, y + 8, width, n2 * 19, s['title2'], 16, INK, caps=False, pitch=19)
        y += 8 + n2 * 19
    if s.get('subtitle'):
        ss = B().capped('подзаголовок', SUB_SIZE)
        n2 = lines(ctx, s['subtitle'], ss, 735)
        sp = ss * SUB_PITCH / SUB_SIZE
        text(slide, X0, y + 8, 735, n2 * sp, s['subtitle'], ss, MUTED, pitch=sp, name='Подзаголовок')
        y += 8 + n2 * sp
    if s.get('footnote'):
        fl = lines(ctx, s['footnote'], 9.5, W)
        text(slide, X0, 495, W, fl * 11.4, s['footnote'], 9.5, NOTE, pitch=11.4, name='Сноска')
    return y


def no_autoshrink(tf):
    """Запрещает PowerPoint ужимать заголовок (normAutofit fontScale): кегль подбирает скрипт, и он един по колоде.
    Иначе после правки текста в PowerPoint заголовок одного слайда становится мельче остальных."""
    bp = tf._txBody.find('{%s}bodyPr' % A)
    for tag in ('normAutofit', 'spAutoFit', 'noAutofit'):
        for el in bp.findall('{%s}%s' % (A, tag)):
            bp.remove(el)
    ext = bp.find('{%s}scene3d' % A)
    na = parse_xml(f'<a:noAutofit xmlns:a="{A}"/>')
    if ext is not None:
        ext.addprevious(na)
    else:
        bp.insert(0, na)


def content_top(slide, s, ctx, where, need, head_bottom):
    """Верх контента, прижатого к низу; предупреждает, если контент наезжает на шапку."""
    top = BOTTOM - need
    if top < head_bottom + 14:
        warn(ctx, where, f'ПЕРЕПОЛНЕНИЕ — контент наезжает на заголовок на {head_bottom + 14 - top:.0f} pt. '
                         'Сократите заголовок или текст')
    return top


def icon(slide, ctx, name, x, y, size, where, white=False, alpha=None):
    if not name:
        return
    if not B().place_icon(slide, ctx['lib'], name, (x, y, size, size), ctx['prs']):
        warn(ctx, where, f'иконки «{name}» нет в библиотеке (см. references/icons.png)')
        return
    grp = slide.shapes[-1]
    if white:  # белая полупрозрачная иконка-водяной знак на градиенте
        for blip in grp._element.iter('{%s}blip' % A):
            for ext in blip.findall('{%s}extLst' % A):
                blip.remove(ext)  # PowerPoint берёт PNG и применяет эффекты
            blip.append(parse_xml(f'<a:lum xmlns:a="{A}" bright="100000"/>'))
            if alpha:
                blip.append(parse_xml(f'<a:alphaModFix xmlns:a="{A}" amt="{int(alpha * 100000)}"/>'))


def picture_fill(slide, ctx, path, x, y, w, h, radius=6, focus=(0.5, 0.4), ellipse=False):
    pic = slide.shapes.add_picture(path, Emu(int(x * E)), Emu(int(y * E)), Emu(int(w * E)), Emu(int(h * E)))
    B().put_image(slide, pic, path, focus=focus)
    geom = pic._element.spPr.find('{%s}prstGeom' % A)
    if ellipse:
        geom.set('prst', 'ellipse')
    elif radius:
        geom.set('prst', 'roundRect')
        av = geom.find('{%s}avLst' % A)
        av.append(parse_xml(f'<a:gd xmlns:a="{A}" name="adj" fmla="val {int(radius / min(w, h) * 100000)}"/>'))
    return pic


def items_of(s):
    return [({'text': g} if isinstance(g, str) else g) for g in (s.get('items') or [])]


# ---------------------------------------------------------------- макеты
def h_cards(slide, s, ctx, where):
    """Серые карточки в ряд (2–4): заголовок сверху, текст прижат к низу. accent: true — градиент."""
    clear_slide(slide)
    hb = header(slide, s, ctx, where)
    its = items_of(s)[:4]
    n = max(1, len(its))
    w = (W - GAP * (n - 1)) / n
    pad = 19
    # высота: заголовок карточки + текст + поле между ними
    body_need = max(lines(ctx, it.get('text', ''), 14, w - 2 * pad) * 17 for it in its) if its else 0
    h = max(s.get('height', 164), 26 + 22 + 40 + body_need + pad)
    top = content_top(slide, s, ctx, where, h, hb)
    for k, it in enumerate(its):
        x = X0 + k * (w + GAP)
        acc = it.get('accent')
        rect(slide, x, top, w, h, grad=GRAD[ctx['dir_no']] if acc else None)
        fg, fg2 = ('FFFFFF', 'FFFFFF') if acc else (INK, TEXT)
        ty = top + 17
        if it.get('icon'):
            icon(slide, ctx, it['icon'], x + pad, ty, 22, where, white=acc)
            ty += 32
        if it.get('title'):
            text(slide, x + pad, ty, w - 2 * pad, 24, it['title'], 18, fg, bold=True)
        if it.get('text'):
            sz, ok = fit_size(ctx, it['text'], 14, w - 2 * pad, h - (ty - top) - 30 - pad, 17 / 14, role='текст карточки')
            if not ok:
                warn(ctx, where, f'ПЕРЕПОЛНЕНИЕ в карточке {k + 1}')
            text(slide, x + pad, top + 40, w - 2 * pad, h - 40 - pad, it['text'], sz, fg2, anchor='bottom',
                 pitch=sz * 17 / 14)


def h_stats(slide, s, ctx, where):
    """Крупные цифры: цифра сверху, подпись прижата к низу. Слова и «/5» в цифре — мелко.
    Акцентная плитка (accent: true, первой) при нескольких рядах занимает левую колонку во всю высоту —
    как «Цели проекта» в эталоне."""
    clear_slide(slide)
    hb = header(slide, s, ctx, where)
    its = items_of(s)[:9]
    acc = its[0] if its and its[0].get('accent') else None
    cols = s.get('columns') or (3 if len(its) != 4 else 2)
    tall = acc is not None and len(its) > cols  # плитка целей во всю высоту
    gcols = cols - 1 if tall else cols
    cells = its[1:] if tall else its
    rows = max(1, (len(cells) + gcols - 1) // gcols)
    w = (W - GAP * (cols - 1)) / cols
    h = s.get('height') or (129 if rows > 1 else 196)
    need = rows * h + (rows - 1) * GAP
    top = content_top(slide, s, ctx, where, need, hb)
    col = ACCENT[ctx['dir_no']]
    pad = 19
    if tall:
        accent_tile(slide, ctx, where, acc, X0, top, w, need, pad, 1)
    gx = X0 + (w + GAP if tall else 0)
    for k, it in enumerate(cells):
        x = gx + (k % gcols) * (w + GAP)
        y = top + (k // gcols) * (h + GAP)
        if it.get('accent'):
            accent_tile(slide, ctx, where, it, x, y, w, h, pad, k + 1)
            continue
        rect(slide, x, y, w, h)
        val = str(it.get('value', ''))
        vs, _ = fit_size(ctx, val, VALUE_SIZE, w - 2 * pad, 60, role='цифра')
        text(slide, x + pad, y + 18, w - 2 * pad, vs * 1.15, val, vs, col, small=(vs * 0.45, SMALL_IN_VALUE),
             pitch=vs * 1.15, name='Цифра')
        if it.get('text'):
            ty = y + 18 + vs * 1.15 + 8
            ts, ok = fit_size(ctx, it['text'], 14, w - 2 * pad, y + h - 16 - ty, 17 / 14, role='подпись цифры')
            if not ok:
                warn(ctx, where, f'ПЕРЕПОЛНЕНИЕ в подписи цифры {k + 1}')
            text(slide, x + pad, ty, w - 2 * pad, y + h - 16 - ty, it['text'], ts, TEXT, anchor='bottom',
                 pitch=ts * 17 / 14, name='Подпись')


def accent_tile(slide, ctx, where, it, x, y, w, h, pad, k):
    """Акцентная плитка с текстом: заголовок сверху, пункты прижаты к низу («Цели проекта»)."""
    rect(slide, x, y, w, h, grad=GRAD[ctx['dir_no']])
    ty = y + 16
    if it.get('title'):
        tn = lines(ctx, it['title'], 15, w - 2 * pad, bold=True)
        text(slide, x + pad, ty, w - 2 * pad, tn * 18, it['title'], 15, 'FFFFFF', bold=True, pitch=18)
        ty += tn * 18 + 14
    pts = it.get('points') or ([it['text']] if it.get('text') else [])
    ps, ok = fit_size(ctx, pts, 12, w - 2 * pad - 14, y + h - 17 - ty, 14.5 / 12, gap=5, role='текст плитки')
    if not ok:
        warn(ctx, where, f'ПЕРЕПОЛНЕНИЕ в плитке {k}')
    text(slide, x + pad, ty, w - 2 * pad, y + h - 17 - ty, pts, ps, 'FFFFFF', pitch=ps * 14.5 / 12,
         bullet='FFFFFF' if len(pts) > 1 else None, gap=5, anchor='bottom')


def h_speakers(slide, s, ctx, where):
    """Спикеры 3 × 2 (эталон): фото 79 × 89 со скруглением, имя 14 pt цветом направления, должность 11 pt."""
    clear_slide(slide)
    hb = header(slide, s, ctx, where)
    its = items_of(s)[:9]
    cols = 3 if len(its) not in (1, 2, 4) else min(2, len(its))
    rows = (len(its) + cols - 1) // cols
    w = (W - GAP * (cols - 1)) / cols
    h = 115
    top = content_top(slide, s, ctx, where, rows * h + (rows - 1) * GAP, hb)
    col = ACCENT[ctx['dir_no']]
    st = s.get('stat')  # общая цифра над карточками: «9,3/10 — средняя оценка спикеров»
    if st:
        vs = B().capped('цифра', VALUE_SIZE)
        vy = top - 18 - vs * 1.15
        if vy < hb + 10:
            warn(ctx, where, 'ПЕРЕПОЛНЕНИЕ — цифра над спикерами наезжает на заголовок')
        val = str(st.get('value', ''))
        vw = ctx['fitter'].width(re.sub(SMALL_IN_VALUE, '', val), vs) + sum(
            ctx['fitter'].width(m, vs * 0.45) for m in re.findall(SMALL_IN_VALUE, val))
        text(slide, X0, vy, vw + 12, vs * 1.15, val, vs, col, small=(vs * 0.45, SMALL_IN_VALUE),
             pitch=vs * 1.15, name='Цифра')
        text(slide, X0 + vw + 24, vy, 220, vs * 1.15, st.get('text', ''), 14, TEXT, anchor='middle', pitch=17)
    pw, ph = 79, 89
    for k, it in enumerate(its):
        x = X0 + (k % cols) * (w + GAP)
        y = top + (k // cols) * (h + GAP)
        rect(slide, x, y, w, h)
        py = y + (h - ph) / 2
        path = B().asset('photos', it.get('photo'))
        if path:
            picture_fill(slide, ctx, path, x + 19, py, pw, ph, radius=6, focus=(0.5, 0.2))
        else:
            rect(slide, x + 19, py, pw, ph, fill='E3E7F0').name = 'Заглушка фото'
            icon(slide, ctx, 'people', x + 19 + pw / 2 - 14, py + ph / 2 - 14, 28, where)
            if it.get('photo'):
                warn(ctx, where, f'фото «{it["photo"]}» не найдено')
            else:
                warn(ctx, where, f'нет фото спикера {k + 1} — стоит заглушка')
        tx = x + 19 + pw + 13
        tw = x + w - 12 - tx
        name = it.get('name', 'Фамилия Имя')
        role = it.get('role', '')
        # имя в одну строку, затем должность — уменьшаются по правилу единого кегля
        ns = B().capped('имя', 14)
        while lines(ctx, name, ns, tw, bold=True) > 1 and ns > 11:
            ns -= 0.5
        rs = B().capped('должность', 11)
        while True:
            nl = lines(ctx, name, ns, tw, bold=True)
            rl = lines(ctx, role, rs, tw) if role else 0
            block = nl * ns * 1.2 + 15 + rl * rs * 1.2
            if block <= h - 28 or rs <= 8.5:
                break
            rs -= 0.5
        if block > h - 28:
            warn(ctx, where, f'ПЕРЕПОЛНЕНИЕ — спикер {k + 1}: сократите должность')
        B().note_size('имя', 14, ns)
        B().note_size('должность', 11, rs)
        text(slide, tx, y + 16, tw, nl * ns * 1.2 + 2, name, ns, col, bold=True, pitch=ns * 1.2, name='Имя')
        if role:
            text(slide, tx, y + 16 + nl * ns * 1.2 + 15, tw, rl * rs * 1.2 + 2, role, rs, TEXT, pitch=rs * 1.2,
                 name='Должность')


def h_features(slide, s, ctx, where):
    """Высокая градиентная карточка слева (заголовок + крупная иконка) и сетка 3 × 2 серых ячеек с иконками."""
    clear_slide(slide)
    hb = header(slide, s, ctx, where)
    its = items_of(s)[:6]
    h = 292
    top = content_top(slide, s, ctx, where, h, hb)
    n = ctx['dir_no']
    lw = 220
    hl = B().asset('cards', f'highlight_{n:02d}')
    picture_fill(slide, ctx, hl, X0, top, lw, h, radius=6, focus=(0.5, 0.5))
    if s.get('card_title'):
        text(slide, X0 + 20, top + 20, lw - 40, 120, s['card_title'], 16, 'FFFFFF', bold=True, pitch=19)
    if s.get('card_icon'):
        icon(slide, ctx, s['card_icon'], X0 + 20, top + h - 76, 56, where, white=True, alpha=0.5)
    cols, rows = 3, 2
    gx = X0 + lw + GAP
    cw = (X1 - gx - GAP * (cols - 1)) / cols
    ch = (h - GAP) / 2
    for k, it in enumerate(its):
        x = gx + (k % cols) * (cw + GAP)
        y = top + (k // cols) * (ch + GAP)
        rect(slide, x, y, cw, ch)
        icon(slide, ctx, it.get('icon'), x + 20, y + 20, 16, where)
        ts, ok = fit_size(ctx, it.get('text', ''), 12.5, cw - 40, ch - 60, 15 / 12.5, role='текст ячейки')
        if not ok:
            warn(ctx, where, f'ПЕРЕПОЛНЕНИЕ в ячейке {k + 1}')
        text(slide, x + 20, y + 58, cw - 40, ch - 70, it.get('text', ''), ts, TEXT, pitch=ts * 15 / 12.5)


def h_rows(slide, s, ctx, where):
    """Строки-форматы: слева градиентная метка с заголовком и иконкой-водяным знаком, справа список в 1–2 колонки."""
    clear_slide(slide)
    hb = header(slide, s, ctx, where)
    its = items_of(s)[:4]
    lw = 262
    pad = 15
    heights = []
    for it in its:
        pts = it.get('points') or []
        ncol = it.get('columns') or (2 if len(pts) > 3 else 1)
        per = (len(pts) + ncol - 1) // ncol
        colw = (X1 - (X0 + lw) - 2 * pad - (ncol - 1) * 20) / ncol
        need = 0
        for c in range(ncol):
            chunk = pts[c * per:(c + 1) * per]
            need = max(need, lines(ctx, chunk, 13, colw - 14) * 14.5 + (len(chunk) - 1) * 9)
        heights.append(max(82, need + pad + 11))  # сверху 15, снизу 11 — как в эталоне
    caps_h = 25  # подпись группы над строкой: «Базовый формат», «Дополнительно» (эталон)
    # подпись группы: над первой строкой 25 pt, между строками — 5 pt после строки и 23 pt до следующей
    total = sum(heights) + GAP * (len(its) - 1) + sum((caps_h if k == 0 else 28 - GAP)
                                                       for k, it in enumerate(its) if it.get('group'))
    top = content_top(slide, s, ctx, where, total, hb)
    y = top
    grad = GRAD[ctx['dir_no']]
    for k, it in enumerate(its):
        h = heights[k]
        if it.get('group'):
            if k:
                y += 5 - GAP
            text(slide, X0, y, 504, 15, it['group'], 12, INK, name='Подпись группы')
            y += caps_h if k == 0 else 23
        rect(slide, X0, y, W, h)
        muted = it.get('muted')  # серая метка вместо градиента (правка автора: «формат взаимодействия»)
        if muted:
            rect(slide, X0, y, lw + 6, h, fill='E3E7F0')
        else:
            rect(slide, X0, y, lw + 6, h, grad=grad)
        # правый край метки прямой: накрываем скругление серой полосой
        cover = rect(slide, X0 + lw, y, 6, h, fill=CARD, radius=0)
        if it.get('icon'):
            icon(slide, ctx, it['icon'], X0 + lw - 96, y + h - 80, 88, where, white=not muted, alpha=0.35)
            if muted:  # синяя иконка тоже полупрозрачная
                for blip in slide.shapes[-1]._element.iter('{%s}blip' % A):
                    for ext in blip.findall('{%s}extLst' % A):
                        blip.remove(ext)
                    blip.append(parse_xml(f'<a:alphaModFix xmlns:a="{A}" amt="30000"/>'))
        text(slide, X0 + 16, y + 14, lw - 110, h - 28, it.get('title', ''), 14,
             ACCENT[ctx['dir_no']] if muted else 'FFFFFF', pitch=17)
        pts = it.get('points') or []
        ncol = it.get('columns') or (2 if len(pts) > 3 else 1)
        per = (len(pts) + ncol - 1) // ncol
        colw = (X1 - (X0 + lw) - 2 * pad - (ncol - 1) * 20) / ncol
        for c in range(ncol):
            chunk = pts[c * per:(c + 1) * per]
            text(slide, X0 + lw + pad + c * (colw + 20), y + pad, colw, h - 2 * pad, chunk, 13, INK, pitch=14.5,
                 bullet=ACCENT[ctx['dir_no']], gap=9)
        y += h + GAP


def h_table(slide, s, ctx, where):
    """Таблица эталона: шапка на градиенте, первая строка выделена, зебра, линии #DCE0EB."""
    clear_slide(slide)
    hb = header(slide, s, ctx, where)
    data = s.get('table') or {}
    cols = data.get('columns') or []
    rows = data.get('rows') or []
    hl = set(data.get('highlight', [0]))
    ncol = max([len(cols)] + [len(r) for r in rows])
    widths = data.get('widths')
    if not widths:
        lens = [max([len(re.sub(r'\n.*', '', str(c))) for c in [cols[i] if i < len(cols) else '']
                     + [r[i] if i < len(r) else '' for r in rows]] + [6]) for i in range(ncol)]
        widths = [max(0.12, l / sum(lens)) for l in lens]
    widths = [w / sum(widths) for w in widths]
    row_h = 37
    nrow = len(rows) + 1
    need = nrow * row_h
    top = hb + 22 if s.get('align') == 'top' else max(hb + 20, BOTTOM - need)
    if top + need > BOTTOM + 8:
        warn(ctx, where, 'ПЕРЕПОЛНЕНИЕ — таблица не помещается, разбейте её')
    gf = slide.shapes.add_table(nrow, ncol, Emu(int(X0 * E)), Emu(int(top * E)), Emu(int(W * E)),
                                Emu(int(need * E)))
    tbl = gf.table
    tblPr = tbl._tbl.find('{%s}tblPr' % A)
    for k in list(tblPr):
        tblPr.remove(k)
    for a in list(tblPr.attrib):
        del tblPr.attrib[a]
    for i, w in enumerate(widths):
        tbl.columns[i].width = Emu(int(W * w * E))
    g = GRAD[ctx['dir_no']]
    col = ACCENT[ctx['dir_no']]
    for r in range(nrow):
        tbl.rows[r].height = Emu(int(row_h * E))
        vals = cols if r == 0 else rows[r - 1]
        for c in range(ncol):
            cell = tbl.cell(r, c)
            v = str(vals[c]) if c < len(vals) else ''
            main, _, note = v.partition('\n')
            cell.margin_left = cell.margin_right = Emu(int(14 * E))
            cell.margin_top = cell.margin_bottom = Emu(int(3 * E))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            run = p.add_run()
            run.text = B().nbsp(main)
            f = run.font
            f.name = FONT
            f.size = Pt(13.5)
            if r == 0:
                f.color.rgb = RGBColor.from_string('FFFFFF')
            elif (r - 1) in hl:
                f.color.rgb = RGBColor.from_string(col)
                f.name = FONT_DEMI
            else:
                f.color.rgb = RGBColor.from_string(INK)
            if note:
                p2 = tf.add_paragraph()
                p2.alignment = p.alignment
                r2 = p2.add_run()
                r2.text = note
                r2.font.size = Pt(7.5)
                r2.font.name = FONT
                r2.font.color.rgb = RGBColor.from_string(INK)
            tcPr = cell._tc.get_or_add_tcPr()
            if r == 0:
                # градиент шапки идёт через всю таблицу: каждой ячейке — свой участок
                x0 = sum(widths[:c])
                x1 = x0 + widths[c]
                c0, c1 = mix(g[0], g[1], x0), mix(g[0], g[1], x1)
                tcPr.append(parse_xml(
                    f'<a:gradFill xmlns:a="{A}" rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="{c0}"/></a:gs>'
                    f'<a:gs pos="100000"><a:srgbClr val="{c1}"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>'))
            else:
                fill = 'EAEDF5' if (r - 1) in hl else ('F4F6FA' if r % 2 == 1 else 'FFFFFF')
                tcPr.append(parse_xml(f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="{fill}"/></a:solidFill>'))
            for side in ('lnL', 'lnR', 'lnT', 'lnB'):
                tcPr.insert(0 if side == 'lnL' else len(tcPr) - 1, parse_xml(
                    f'<a:{side} xmlns:a="{A}" w="9525"><a:solidFill><a:srgbClr val="DCE0EB"/></a:solidFill></a:{side}>'))


def _img_path(pic):
    """Сохраняет картинку фигуры во временный файл (для повторного кадрирования)."""
    import tempfile
    fd, path = tempfile.mkstemp(suffix='.' + pic.image.ext)
    with os.fdopen(fd, 'wb') as f:
        f.write(pic.image.blob)
    return path


def mix(a, b, t):
    ca = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    cb = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return ''.join(f'{round(x + (y - x) * t):02X}' for x, y in zip(ca, cb))


def h_course(slide, s, ctx, where):
    """Страница курса (эталон): название, описание 12 pt во всю ширину, «Продолжительность / 24 ак. ч»,
    внизу две карточки — «Содержание курса» (нумерованный список «**Тема.** описание») и «Особенности»."""
    clear_slide(slide)
    s2 = dict(s)
    s2.pop('subtitle', None)
    hb = header(slide, s2, ctx, where, max_lines=2)
    lead = s.get('lead') or s.get('subtitle') or ''
    y = hb + 6
    if lead:
        ls, _ = fit_size(ctx, lead, 12, W, 60, 14.5 / 12, role='описание курса')
        nl = lines(ctx, lead, ls, W)
        text(slide, X0, y, W, nl * ls * 14.5 / 12, lead, ls, INK, pitch=ls * 14.5 / 12, name='Описание')
        y += nl * ls * 14.5 / 12
    if s.get('duration'):
        tb = text(slide, X0, y + 10, 310, 34, ['Продолжительность', '**' + s['duration'] + '**'], 14, INK,
                  pitch=17, name='Продолжительность')
        r0 = tb.text_frame.paragraphs[0].runs[0].font
        r0.size = Pt(11)
        r0.color.rgb = RGBColor.from_string(NOTE)
        y += 10 + 34
    top = y + 22
    mods = s.get('modules') or []
    feats = s.get('features') or []
    lw = 500 if feats else W
    rw = W - lw - 13
    h = BOTTOM - top
    pad = 20
    rect(slide, X0, top, lw, h)
    text(slide, X0 + pad, top + 16, lw - 2 * pad, 18, s.get('modules_title') or 'Содержание курса', 14, INK,
         bold=True, name='Подзаголовок карточки')
    ms, ok = fit_size(ctx, mods, 12, lw - 2 * pad - 16, h - 51 - 14, 14.5 / 12, gap=4, role='модули')
    if not ok:
        warn(ctx, where, 'ПЕРЕПОЛНЕНИЕ — модули не помещаются, сократите описания')
    text(slide, X0 + pad, top + 51, lw - 2 * pad, h - 51 - 14, mods, ms, INK, pitch=ms * 14.5 / 12, gap=4,
         numbered=True, name='Модули')
    if feats:
        rx = X0 + lw + 13
        rect(slide, rx, top, rw, h)
        text(slide, rx + 17, top + 16, rw - 34, 18, s.get('features_title') or 'Особенности', 14, INK, bold=True,
             name='Подзаголовок карточки')
        fs, ok = fit_size(ctx, feats, 12, rw - 34 - 14, h - 51 - 14, 14.5 / 12, gap=4, role='особенности')
        if not ok:
            warn(ctx, where, 'ПЕРЕПОЛНЕНИЕ — особенности не помещаются')
        text(slide, rx + 17, top + 51, rw - 34, h - 51 - 14, feats, fs, INK, pitch=fs * 14.5 / 12, gap=4,
             bullet=INK, name='Особенности')


def h_team(slide, s, ctx, where):
    """Контакты эталона: карточки во всю ширину, круглое фото, ФИО прописными, должность, почта."""
    clear_slide(slide)
    s2 = dict(s)
    s2.setdefault('title', 'Контакты')
    header(slide, s2, ctx, where, size=40, max_lines=1)
    its = items_of(s)[:3]
    n = len(its)
    top0 = 108
    h = (BOTTOM + 12 - top0 - GAP * (n - 1)) / max(1, n)
    card = B().asset('cards', f'card_{ctx["dir_no"]:02d}')
    light = ctx['dir_no'] in (5, 6)
    for k, it in enumerate(its):
        y = top0 + k * (h + GAP)
        first = k == 0
        if first:
            picture_fill(slide, ctx, card, X0, y, W, h, radius=6)
        else:
            rect(slide, X0, y, W, h, fill='EAEDF5')
        fg = ('FFFFFF' if not light else '0E2569') if first else INK
        fg2 = ('FFFFFF' if not light else '0E2569') if first else MUTED
        ps = min(84, h - 40)
        path = B().asset('photos', it.get('photo'))
        if path:
            picture_fill(slide, ctx, path, X0 + 20, y + 20, ps, ps, ellipse=True, focus=(0.5, 0.2))
        tx = X0 + 20 + (ps + 20 if path else 0)
        text(slide, tx, y + 28, W - (tx - X0) - 20, 34, it.get('name', 'Фамилия Имя'), 28, fg, caps=False, pitch=32)
        if it.get('role'):
            text(slide, tx, y + 62, W - (tx - X0) - 20, 24, it['role'], 18, fg2, pitch=21)
        contacts = [c for c in (it.get('email'), it.get('phone')) if c]
        for j, c in enumerate(contacts):
            ix = X0 + 20 + j * 300
            icon(slide, ctx, 'mail' if '@' in c else 'phone', ix, y + h - 42, 24, where, white=first and not light)
            text(slide, ix + 34, y + h - 40, 260, 22, c, 16, fg if first else ACCENT[ctx['dir_no']])
        if it.get('photo') and not path:
            warn(ctx, where, f'фото контакта «{it["photo"]}» не найдено')


def h_title(slide, s, ctx, where):
    """Обложка эталона: белое поле, диагональная панель с портретом, название 28 pt до 3 строк (раздел — 48 pt),
    внизу слева синяя черта, под ней пометка (синим) и дата."""
    title = s.get('title', '')
    chapter = s.get('layout') == 'chapter'
    t = next((p for p in slide.placeholders if p.placeholder_format.type in (1, 3)), None)
    # портрет — справа во всю высоту и под текстом
    for pic in [x for x in slide.shapes if x.name.startswith('Image')]:
        pic.left, pic.top, pic.width, pic.height = (Emu(int(v * E)) for v in (480, 0, 480, 540))
        B().put_image(slide, pic, _img_path(pic), focus=(0.5, 0.0))
        tree = slide.shapes._spTree
        tree.remove(pic._element)
        tree.insert(2, pic._element)
    size0 = 48 if chapter else 28
    tw = 428 if chapter else 445  # правее начинается портрет: текст не должен на него заходить
    sz, ok = fit_size(ctx, title, size0, tw, (2 if chapter else 3) * size0 * 1.07, 1.07, floor=20, caps=False)
    if not ok:
        warn(ctx, where, 'ПЕРЕПОЛНЕНИЕ — название не помещается')
    n = lines(ctx, title, sz, tw, caps=False)
    h = n * sz * 1.07
    centre = 270
    if t is not None:
        t.left, t.top, t.width, t.height = (Emu(int(v * E)) for v in (29 if chapter else 35, centre - h / 2, tw, h + 6))
        tf = t.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.text = B().nbsp(title)
        for p in tf.paragraphs:
            p.line_spacing = Pt(sz * 1.07)
            for r in p.runs:
                r.font.size = Pt(sz)
                r.font.color.rgb = RGBColor.from_string(ACCENT[ctx['dir_no']])
    for sh in list(slide.placeholders):  # нижние поля прототипа заменяем своими
        if sh.placeholder_format.type not in (1, 3, 13):
            sh._element.getparent().remove(sh._element)
    if s.get('kicker'):
        kn = lines(ctx, s['kicker'], 20, 420, caps=False)
        ty = t.top / E if t is not None else 200
        text(slide, X0, ty - kn * 22 - 8, 420, kn * 22, s['kicker'], 20, 'ACB6C3', caps=False, pitch=22)
    note = [x for x in (s.get('note'), s.get('date')) if x]
    if note:
        col = ACCENT[ctx['dir_no']]
        tb = text(slide, 41, 473, 330, 34, note, 12.5, INK, pitch=15)  # вровень с синей чертой (правка автора)
        tb.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string(col)
        if len(note) > 1 and s.get('date_muted'):
            tb.text_frame.paragraphs[1].runs[0].font.color.rgb = RGBColor.from_string('ACB6C3')


HOUSE = {
    'cards': h_cards, 'stats': h_stats, 'speakers': h_speakers, 'features': h_features, 'rows': h_rows,
    'datatable': h_table, 'course': h_course, 'team': h_team, 'title': h_title, 'chapter': h_title,
}
