#!/usr/bin/env python3
"""Ревьюер презентации: проверяет PPTX перед отдачей и исправляет висячие предлоги.

    python3 review.py deck.pptx                 # отчёт по слайдам; код выхода 1, если есть ошибки
    python3 review.py deck.pptx --fix out.pptx  # то же + неразрывные пробелы после предлогов и союзов
    python3 review.py deck.pptx --json          # отчёт в JSON

Проверки (ОШИБКА — исправить до отдачи, ВНИМАНИЕ — посмотреть глазами):
- текст: не помещается в рамку или карточку, налезает на другой текст, иконку или край фото; слово шире рамки;
- поля: контент левее 30 pt, правее 930 pt, ниже 487 pt; края почти совпадают (1–3 pt) — не выровнены;
- фото: растянуто; портрет-вырезка обрезан рамкой или краем слайда (плечо, голова); фото вылезает из карточки;
- иконки: искажены пропорции, разный размер в ряду, сдвинуты относительно соседних, прижаты к краю карточки,
  в одной из одинаковых карточек нет иконки;
- шрифты: только VTB Omega UI / Medium (и в теме); капс (cap="all", слова прописными);
- кегль: заголовки разного кегля или цвета, автосжатие PowerPoint (normAutofit fontScale), разный кегль у
  одинаковых элементов (в ряду и по колоде), мельче 9 pt (заголовок мельче 20 pt) — просить сократить;
- заголовок длиннее 3 строк; пустые плейсхолдеры и поля; фигуры за пределами слайда;
- висячие предлоги и союзы (обычный пробел после «в», «и», «на», «для»…) — исправляются с --fix;
- заглушки (XX, «Фамилия Имя», TBD, ПРОВЕРИТЬ…) и звёздочки без сноски — список для автора.
"""
import argparse, collections, io, json, os, re, sys

from PIL import Image
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_deck import Fitter, E, eff_size, xp, nbsp, stype  # noqa: E402

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
FONTS = {'VTB Omega UI', 'VTB Omega UI Medium'}
LEFT, RIGHT, BOTTOM = 30, 930, 487  # поля и низ контента фирменного стиля
GRID_BOTTOM = 500                     # низ сетки мастер-шаблона
TITLE_MAX_LINES = 3
ICON_MAX = 40          # фигура меньше — иконка
DECOR_MAX = 100        # группа меньше — цельный значок (иконка, водяной знак): внутрь не заходим
ACRONYMS = {'ВТБ', 'МСФО', 'ДПО', 'АНО', 'ДО', 'РФ', 'LMS', 'IFRS', 'IAS', 'ACCA', 'CIMA', 'ЦФА', 'КВ', 'АХР',
            'ДВОПиИ', 'НМА', 'НИОКР', 'KPI', 'ECL', 'FVTPL', 'FVOCI', 'CoR', 'TBD', 'ПАО', 'РОСТ', 'ШОиФД', 'ВШМ',
            'СПбГУ', 'ОАЭ', 'HR', 'IT', 'B2B', 'ESG', 'ИИ', 'AI', 'MBA', 'EMBA', 'QR', 'PDF', 'PPTX', 'TBC'}
STUBS = re.compile(r'(?<![\w])(X{2,}|X,X|[xх]{2,}|[xх],[xх])(?![\w])|Фамилия Имя|Образец текста|\bTBD\b|ПРОВЕРИТЬ|'
                   r'ЗАГЛУШК|[Ll]orem', re.I)
COVER_LAYOUTS = re.compile(r'Обложк|Раздел|Титул|Section|Title Slide', re.I)


# ---------------------------------------------------------------- геометрия
def walk(shapes, tf=None):
    """Все фигуры с абсолютными координатами (pt) с учётом групп: (фигура, x, y, w, h, группа)."""
    tf = tf or (lambda x, y, w, h: (x, y, w, h))
    for sh in shapes:
        if sh.width is None:
            continue
        box = tf(sh.left / E, sh.top / E, sh.width / E, sh.height / E)
        yield sh, box
        if stype(sh) == 6 and max(box[2], box[3]) > DECOR_MAX:  # группа-контейнер: координаты детей → абсолютные
            xfrm = sh._element.find('{%s}grpSpPr/{%s}xfrm' % (P, A))
            off, ext = xfrm.find('{%s}off' % A), xfrm.find('{%s}ext' % A)
            cho, che = xfrm.find('{%s}chOff' % A), xfrm.find('{%s}chExt' % A)
            if che is None or int(che.get('cx')) == 0 or int(che.get('cy')) == 0:
                continue
            gx, gy, gw, gh = box
            kx = gw / (int(che.get('cx')) / E)
            ky = gh / (int(che.get('cy')) / E)
            cx, cy = int(cho.get('x')) / E, int(cho.get('y')) / E

            def child(x, y, w, h, gx=gx, gy=gy, kx=kx, ky=ky, cx=cx, cy=cy):
                return gx + (x - cx) * kx, gy + (y - cy) * ky, w * kx, h * ky
            for item in walk(sh.shapes, child):
                yield item[0], item[1]


def inter(a, b):
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return (w, h) if w > 0 and h > 0 else (0, 0)


def inside(a, b, tol=1):
    return a[0] >= b[0] - tol and a[1] >= b[1] - tol and a[0] + a[2] <= b[0] + b[2] + tol and \
        a[1] + a[3] <= b[1] + b[3] + tol


def centre_in(a, b):
    cx, cy = a[0] + a[2] / 2, a[1] + a[3] / 2
    return b[0] <= cx <= b[0] + b[2] and b[1] <= cy <= b[1] + b[3]


# ---------------------------------------------------------------- текст
def para_size(p, sh, slide):
    sizes = [r.font.size.pt for r in p.runs if r.font.size and r.text.strip()]
    if sizes:
        return max(sizes)
    epr = p._p.find('{%s}endParaRPr' % A)
    if epr is not None and epr.get('sz'):
        return int(epr.get('sz')) / 100
    return eff_size(sh, slide) if sh.is_placeholder else 18.0


def font_scale(sh):
    na = xp(sh._element, './/a:bodyPr/a:normAutofit')
    if na and na[0].get('fontScale'):
        return int(na[0].get('fontScale')) / 100000
    return 1.0


def chain(sh, slide):
    """Элементы, от которых фигура наследует свойства: сама → плейсхолдер макета → плейсхолдер мастера."""
    out = [sh._element]
    if sh.is_placeholder and slide is not None:
        pf = sh.placeholder_format
        lay = slide.slide_layout
        lp = next((p for p in lay.placeholders if p.placeholder_format.idx == pf.idx), None)
        if lp is not None:
            out.append(lp._element)
        typ = pf.type
        mp = next((p for p in lay.slide_master.placeholders if p.placeholder_format.type == typ), None)
        if mp is not None:
            out.append(mp._element)
    return out


def inherited(sh, slide, q):
    for el in chain(sh, slide):
        r = xp(el, q)
        if r:
            return r[0]
    return None


def body_insets(sh, slide=None):
    g = (lambda k, d: int(v) / E if (v := inherited(sh, slide, './/a:bodyPr/@' + k)) is not None else d)
    return g('lIns', 7.2), g('tIns', 3.6), g('rIns', 7.2), g('bIns', 3.6)


def wrap_widths(fit, text, size, width, bold):
    """Ширины строк после переноса по словам (как Fitter.lines)."""
    out = []
    for para in text.split('\n'):
        cur = ''
        for w in para.split(' '):
            t = (cur + ' ' + w).strip()
            if fit.width(t, size, bold) <= width or not cur:
                cur = t
            else:
                out.append(fit.width(cur, size, bold))
                cur = w
        out.append(fit.width(cur, size, bold))
    return out


def text_extent(sh, box, slide, fit):
    """Где реально стоит текст: (x, y, w, h), высота текста, ширина самого длинного слова, ширина рамки."""
    l, t, r, b = body_insets(sh, slide)
    nowrap = inherited(sh, slide, './/a:bodyPr/@wrap') == 'none'
    inh_pct = inherited(sh, slide, './/a:lvl1pPr/a:lnSpc/a:spcPct/@val')
    x, y, w, h = box
    iw = w - l - r
    k = font_scale(sh)
    total, maxw, widest = 0, 0, 0
    for i, p in enumerate(sh.text_frame.paragraphs):
        txt = ''.join(r_.text for r_ in p.runs)
        size = para_size(p, sh, slide) * k
        bold = any((r_.font.name or '').endswith(('Medium', 'DemiBold', 'Bold')) for r_ in p.runs if r_.text.strip())
        ls = p.line_spacing
        if ls is None and inh_pct:
            pitch = size * 1.2 * int(inh_pct) / 100000
        elif ls is None:
            pitch = size * 1.2
        elif isinstance(ls, float):
            pitch = size * 1.2 * ls
        else:
            pitch = ls.pt * (k if k < 1 else 1)
        indent = 0
        ppr = p._p.pPr
        if ppr is not None and ppr.get('marL'):
            indent = int(ppr.get('marL')) / E
        ws = wrap_widths(fit, txt, size, 1e9 if nowrap else max(1, iw - indent), bold) if txt.strip() else [0]
        sb = p.space_before.pt if (i and p.space_before is not None) else 0
        total += len(ws) * pitch + sb
        if txt.strip():
            maxw = max(maxw, min(iw, max(ws) + indent))
            # самое длинное слово — по кеглю своего фрагмента; PowerPoint переносит и по дефису
            for r_ in p.runs:
                rs = (r_.font.size.pt if r_.font.size else size / k) * k
                rb = (r_.font.name or '').endswith(('Medium', 'DemiBold', 'Bold'))
                for wd in re.split(r'[\s\u00a0]+|(?<=-)', r_.text):
                    if wd:
                        widest = max(widest, fit.width(wd, rs, rb) + indent)
    anchor = inherited(sh, slide, './/a:bodyPr/@anchor') or 't'
    if nowrap:
        widest = 0  # без переноса слово не разорвётся
    algn = (xp(sh._element, './/a:pPr/@algn') or ['l'])[0]
    if anchor == 'b':
        ty = y + h - b - total
    elif anchor == 'ctr':
        ty = y + (h - total) / 2
    else:
        ty = y + t
    tx = x + l + (iw - maxw if algn == 'r' else (iw - maxw) / 2 if algn == 'ctr' else 0)
    return (tx, ty, maxw, total), total, widest, iw, h - t - b


# ---------------------------------------------------------------- висячие предлоги
def fix_hanging(prs):
    """Обычный пробел после предлогов, союзов, цифр и перед тире → неразрывный. Возвращает {слайд: число}."""
    out = collections.Counter()
    for no, slide in enumerate(prs.slides, 1):
        for sh, _ in walk(slide.shapes):
            frames = []
            if sh.has_text_frame:
                frames.append(sh.text_frame)
            if getattr(sh, 'has_table', False) and sh.has_table:
                frames += [c.text_frame for row in sh.table.rows for c in row.cells]
            for tf in frames:
                for p in tf.paragraphs:
                    runs = list(p.runs)
                    src = ''.join(r.text for r in runs)
                    dst = nbsp(src)
                    if dst == src or len(dst) != len(src):
                        continue
                    out[no] += sum(1 for a, b in zip(src, dst) if a != b)
                    pos = 0
                    for r in runs:
                        n = len(r.text)
                        r.text = dst[pos:pos + n]
                        pos += n
    return out


def count_hanging(prs):
    out = collections.Counter()
    for no, slide in enumerate(prs.slides, 1):
        for sh, _ in walk(slide.shapes):
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    src = ''.join(r.text for r in p.runs)
                    out[no] += sum(1 for a, b in zip(src, nbsp(src)) if a != b)
    return out


# ---------------------------------------------------------------- картинки
def image_of(pic):
    try:
        return Image.open(io.BytesIO(pic.image.blob))
    except Exception:  # SVG и прочее — не проверяем пиксели
        return None


def src_rect(pic):
    return pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom


def is_photo(pic):
    """Фото (портрет, сцена), а не векторная графика, иллюстрация, градиент или QR: много цветов и деталей."""
    if hasattr(pic, '_rv_photo'):
        return pic._rv_photo
    from PIL import ImageFilter, ImageStat
    im = image_of(pic)
    res = False
    if im is not None:
        rgba = im.convert('RGBA').resize((96, 96))
        a = rgba.getchannel('A')
        rgb = rgba.convert('RGB')
        mask = a.point(lambda v: 255 if v > 128 else 0)
        if mask.getbbox() and min(im.size) >= 64:
            px, al = rgb.tobytes(), a.tobytes()
            colours = len({px[i * 3:i * 3 + 3] for i in range(len(al)) if al[i] > 128})
            edges = ImageStat.Stat(rgb.convert('L').filter(ImageFilter.FIND_EDGES), mask=mask).mean[0]
            # фото: тысячи оттенков и мелкие детали; градиенты, схемы, полоски и QR — нет
            res = colours > 1000 and edges > 15
    pic._rv_photo = res
    return res


def cut_edges(pic, box, sw, sh_):
    """Стороны, где рамка или край слайда режут непрозрачный объект вырезки (плечо, голову).
    Низ портрета у нижнего края слайда — норма."""
    im = image_of(pic)
    if im is None or im.mode not in ('RGBA', 'LA', 'P'):
        return []
    a = im.convert('RGBA').getchannel('A')
    if a.getextrema()[0] > 200:  # непрозрачная картинка — не вырезка
        return []
    W, H = a.size
    cl, ct, cr, cb = src_rect(pic)
    x, y, w, h = box
    # видимая часть рамки (внутри слайда) → координаты в пикселях исходника
    vx0, vy0 = max(x, 0), max(y, 0)
    vx1, vy1 = min(x + w, sw), min(y + h, sh_)
    if vx1 <= vx0 or vy1 <= vy0:
        return []

    def px(sx, sy):
        u = cl + (sx - x) / w * (1 - cl - cr)
        v = ct + (sy - y) / h * (1 - ct - cb)
        return min(W - 1, max(0, int(u * W))), min(H - 1, max(0, int(v * H)))
    bad = []
    sides = {
        'слева': ((vx0, vy0), (vx0, vy1), vx0 > x + 0.5 or cl > 0.001, vx0 > 1),
        'справа': ((vx1, vy0), (vx1, vy1), vx1 < x + w - 0.5 or cr > 0.001, vx1 < sw - 1),
        'сверху': ((vx0, vy0), (vx1, vy0), vy0 > y + 0.5 or ct > 0.001, vy0 > 1),
        'снизу': ((vx0, vy1), (vx1, vy1), vy1 < y + h - 0.5 or cb > 0.001, vy1 < sh_ - 1),
    }
    for side, (p0, p1, cropped, inner) in sides.items():
        if side == 'снизу' and not inner:
            continue  # бюст, обрезанный нижним краем слайда, — норма
        n, opaque = 40, 0
        for i in range(n):
            t = (i + 0.5) / n
            sx, sy = p0[0] + (p1[0] - p0[0]) * t, p0[1] + (p1[1] - p0[1]) * t
            # точку берём чуть внутри видимой области
            sx += {'слева': 1, 'справа': -1}.get(side, 0)
            sy += {'сверху': 1, 'снизу': -1}.get(side, 0)
            if a.getpixel(px(sx, sy)) > 128:
                opaque += 1
        if opaque >= 2 and (cropped or inner):
            bad.append((side, opaque / n, 'краем картинки посреди слайда' if not cropped else
                        'рамкой' if inner else 'краем слайда'))
    return bad


def opaque_box(pic, box):
    """Для вырезки (PNG с прозрачностью) — рамка непрозрачной части: текст над прозрачным фоном не мешает."""
    im = image_of(pic)
    if im is None or im.mode not in ('RGBA', 'LA', 'P'):
        return box
    bb = im.convert('RGBA').getchannel('A').point(lambda v: 255 if v > 128 else 0).getbbox()
    if not bb:
        return box
    W, H = im.size
    cl, ct, cr, cb = src_rect(pic)
    x, y, w, h = box
    sx = lambda u: x + (u / W - cl) / max(1e-6, 1 - cl - cr) * w
    sy = lambda v: y + (v / H - ct) / max(1e-6, 1 - ct - cb) * h
    x0, y0, x1, y1 = max(x, sx(bb[0])), max(y, sy(bb[1])), min(x + w, sx(bb[2])), min(y + h, sy(bb[3]))
    return (x0, y0, max(0, x1 - x0), max(0, y1 - y0))


def covers(pic, box, rect, level=128):
    """Есть ли под текстом непрозрачная часть картинки (у вырезки прозрачный фон тексту не мешает)."""
    im = image_of(pic)
    if im is None or im.mode not in ('RGBA', 'LA', 'P'):
        return True
    a = im.convert('RGBA').getchannel('A')
    W, H = a.size
    cl, ct, cr, cb = src_rect(pic)
    x, y, w, h = box
    x0, y0 = max(rect[0], x), max(rect[1], y)
    x1, y1 = min(rect[0] + rect[2], x + w), min(rect[1] + rect[3], y + h)
    hits = n = 0
    for i in range(12):
        for j in range(6):
            sx, sy = x0 + (x1 - x0) * (i + 0.5) / 12, y0 + (y1 - y0) * (j + 0.5) / 6
            u = cl + (sx - x) / w * (1 - cl - cr)
            v = ct + (sy - y) / h * (1 - ct - cb)
            n += 1
            hits += a.getpixel((min(W - 1, int(u * W)), min(H - 1, int(v * H)))) > level
    return hits > 0


def aspect_error(pic, box):
    im = image_of(pic)
    if im is None:
        return 0
    W, H = im.size
    cl, ct, cr, cb = src_rect(pic)
    vis = (W * (1 - cl - cr)) / max(1e-6, H * (1 - ct - cb))
    frame = box[2] / max(1e-6, box[3])
    return abs(vis / frame - 1)


def group_distortion(sh):
    xfrm = sh._element.find('{%s}grpSpPr/{%s}xfrm' % (P, A))
    ext, che = xfrm.find('{%s}ext' % A), xfrm.find('{%s}chExt' % A)
    if che is None or not int(che.get('cy')) or not int(ext.get('cy')):
        return 0
    return abs((int(ext.get('cx')) / int(ext.get('cy'))) / (int(che.get('cx')) / int(che.get('cy'))) - 1)


def stroke_ratio(sh, box):
    """Толщина самой толстой обводки внутри группы-иконки относительно её размера.
    Обводка в группе не масштабируется вместе с группой: у уменьшенной иконки линии сливаются."""
    ws = [int(w) / E for w in xp(sh._element, './/a:ln/@w')]
    if not ws:
        return 0
    return max(ws) / max(1, min(box[2], box[3]))


def has_fill(sh):
    sp = sh._element.find('{%s}spPr' % P)
    return sp is not None and (sp.find('{%s}solidFill' % A) is not None or sp.find('{%s}gradFill' % A) is not None
                               or sp.find('{%s}blipFill' % A) is not None)


def transparent(sh):
    return bool(xp(sh._element, './/a:alphaModFix'))


# ---------------------------------------------------------------- ревью
class Review:
    def __init__(self, prs):
        self.prs = prs
        self.fit = Fitter()
        self.items = []
        self.sw, self.sh = prs.slide_width / E, prs.slide_height / E
        self.house = True  # слайд фирменного стиля (макет «Таблица» — шапка, логотип, номер)

    def add(self, slide, level, kind, msg):
        self.items.append(dict(slide=slide, level=level, kind=kind, msg=msg))

    def run(self):
        self.theme_fonts()
        roles = collections.defaultdict(list)    # роль → [(слайд, кегль)]
        titles = []
        for no, slide in enumerate(self.prs.slides, 1):
            self.slide(no, slide, roles, titles)
        self.deck_uniform(roles, titles)
        return self

    # -- тема
    def theme_fonts(self):
        seen = set()
        for m in self.prs.slide_masters:
            for rel in m.part.rels.values():
                if rel.reltype.endswith('/theme') and rel.target_part.partname not in seen:
                    seen.add(rel.target_part.partname)
                    blob = rel.target_part.blob.decode('utf8', 'ignore')
                    for tag in ('majorFont', 'minorFont'):
                        m_ = re.search(r'<a:%s><a:latin typeface="([^"]*)"' % tag, blob)
                        if m_ and m_.group(1) not in FONTS:
                            self.add(0, 'ОШИБКА', 'шрифт', f'в теме {rel.target_part.partname} {tag} = «{m_.group(1)}», '
                                                           'нужен VTB Omega UI (tools/restyle_pptx.py)')

    # -- слайд
    def slide(self, no, slide, roles, titles):
        cover = bool(COVER_LAYOUTS.search(slide.slide_layout.name or ''))
        self.house = 'Таблица' in (slide.slide_layout.name or '')
        shapes = list(walk(slide.shapes))
        texts, icons, photos, cards, decor = [], [], [], [], []
        has_footnote = False
        marks = False
        for sh, box in shapes:
            x, y, w, h = box
            st = stype(sh)
            # за пределами слайда
            if x >= self.sw or y >= self.sh or x + w <= 0 or y + h <= 0:
                self.add(no, 'ОШИБКА', 'поля', f'«{sh.name}» целиком за пределами слайда — удалить')
                continue
            if st == 6:
                if max(w, h) <= DECOR_MAX:
                    thick = stroke_ratio(sh, box)
                    if thick > 0.12:
                        self.add(no, 'ОШИБКА', 'иконка', f'иконка «{sh.name}» с непропорционально толстой обводкой '
                                                         f'({thick:.0%} размера) — линии слились, иконка кривая; '
                                                         'обводку уменьшить вместе с иконкой')
                if max(w, h) <= ICON_MAX:
                    icons.append((sh, box))
                elif max(w, h) <= DECOR_MAX:
                    decor.append((sh, box))
                    if group_distortion(sh) > 0.04:
                        self.add(no, 'ОШИБКА', 'иконка', f'значок «{sh.name}» искажён: пропорции сбиты на {group_distortion(sh):.0%}')
                continue
            if st == 13 and sh.name.startswith('Водяной знак'):  # обрезан по метке сборкой
                decor.append((sh, box))
                continue
            if st == 13:  # картинка
                if max(w, h) <= ICON_MAX:
                    icons.append((sh, box))
                else:
                    photos.append((sh, box))
                continue
            if st in (1, 14, 17) or sh.is_placeholder:
                if not sh.is_placeholder and st == 1 and has_fill(sh) and w > 64 and h > 64:
                    cards.append((sh, box))
                if sh.has_text_frame:
                    texts.append((sh, box))
            if getattr(sh, 'has_table', False) and sh.has_table:
                self.margins(no, sh, box)
        for sh, box in texts:
            txt = sh.text_frame.text
            ph = sh.placeholder_format.type if sh.is_placeholder else None
            if not txt.strip():
                if ph in (13, 16, 15):  # номер, дата, нижний колонтитул
                    continue
                if sh.is_placeholder:
                    self.add(no, 'ОШИБКА', 'пусто', f'пустой плейсхолдер «{sh.name}» — в PowerPoint покажет образец текста; '
                                                    'удалить или заполнить')
                elif not has_fill(sh):
                    self.add(no, 'ОШИБКА', 'пусто', f'пустое текстовое поле «{sh.name}» ({box[0]:.0f}, {box[1]:.0f}) — удалить')
                continue
            if ph == 13:
                continue
            self.text_checks(no, slide, sh, box, ph, cover, roles, titles)
            if re.match(r'^\s*(\*|¹|\d\s?[—–-])', txt) and box[1] > 470:
                has_footnote = True
            elif re.search(r'\*|¹|²', txt):
                marks = True
            for m in STUBS.finditer(txt):
                self.add(no, 'ВНИМАНИЕ', 'заглушка', f'«{m.group(0)}» в «{txt[:40]}…» — данные от автора')
                break
        if marks and not has_footnote:
            self.add(no, 'ВНИМАНИЕ', 'сноска', 'есть «*» или «¹», но нет сноски внизу слайда')
        if slide.has_notes_slide:
            nt = slide.notes_slide.notes_text_frame.text if slide.notes_slide.notes_text_frame else ''
            if re.search(r'ПРОВЕРИТЬ|ЗАГЛУШК', nt):
                self.add(no, 'ВНИМАНИЕ', 'заглушка', f'в заметках: {nt[:120]}')
        self._texts = texts
        self.overlaps(no, slide, texts, icons, photos, decor)
        self.picture_checks(no, photos, cards)
        self.icon_checks(no, icons, cards)
        for d, box in decor:  # водяные знаки и крупные значки — внутри своей карточки
            # карточка, с которой значок пересекается больше всего (у метки строки — сама строка)
            over = [(inter(box, c[1])[0] * inter(box, c[1])[1], -c[1][2] * c[1][3], c) for c in cards]
            over = [o for o in over if o[0] > 0]
            card = max(over, key=lambda o: o[:2])[2] if over else None
            if card is not None and not inside(box, card[1], 1):
                self.add(no, 'ОШИБКА', 'иконка', f'значок «{d.name}» вылезает за карточку')
        self.alignment(no, [s for s in texts if not s[0].is_placeholder and s[0].text_frame.text.strip()], 'текстов')
        self.alignment(no, cards + photos, 'карточек и фото')
        for sh, box in cards + photos + [i for i in icons]:
            self.margins(no, sh, box)

    def text_checks(self, no, slide, sh, box, ph, cover, roles, titles):
        ext, need, widest, iw, ih = text_extent(sh, box, slide, self.fit)
        sh._rv_ext = ext
        is_title = ph in (1, 3)
        k = font_scale(sh)
        if k < 1:
            self.add(no, 'ОШИБКА', 'кегль', f'«{sh.text_frame.text[:30]}…»: PowerPoint сжал текст до {k:.0%} (normAutofit) — '
                                             'выключить автоподбор и выставить кегль, единый по колоде')
        sizes = collections.Counter()
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                s = (r.font.size.pt if r.font.size else para_size(p, sh, slide)) * k
                sizes[s] += len(r.text)
                fn = r.font.name
                if fn and not fn.startswith('+') and fn not in FONTS:
                    self.add(no, 'ОШИБКА', 'шрифт', f'«{r.text[:30]}» набран «{fn}» — нужен VTB Omega UI')
                rpr = r._r.find('{%s}rPr' % A)
                if rpr is not None and rpr.get('cap') in ('all', 'small'):
                    self.add(no, 'ОШИБКА', 'капс', f'«{r.text[:30]}»: включены прописные (cap) — выключить')
                for wd in re.findall(r'[A-ZА-ЯЁ]{4,}[A-ZА-ЯЁ-]*', r.text):
                    if wd not in ACRONYMS and len(wd) > 4:
                        self.add(no, 'ОШИБКА', 'капс', f'слово прописными «{wd}» — набрать обычным регистром')
        if not sizes:
            return
        main = sizes.most_common(1)[0][0]
        small = min(sizes)
        if is_title and not cover:
            n = self.fit.lines(sh.text_frame.text, main, iw)
            titles.append((no, main, self.title_colour(sh)))
            if n > TITLE_MAX_LINES:
                self.add(no, 'ОШИБКА', 'заголовок', f'заголовок {n} строк при пределе {TITLE_MAX_LINES} — сократить')
            if main < 20:
                self.add(no, 'ОШИБКА', 'кегль', f'заголовок {main:g} pt — слишком мелко, попросить автора сократить')
            if not xp(sh._element, './/a:bodyPr/a:noAutofit'):
                self.add(no, 'ВНИМАНИЕ', 'кегль', 'в заголовке не выключен автоподбор: после правки PowerPoint '
                                                  'может сжать его мельче остальных (нужен noAutofit)')
        elif small < 9 and box[1] < 470:
            self.add(no, 'ОШИБКА', 'кегль', f'«{sh.text_frame.text[:30]}…» {small:g} pt — мельче 9 pt не читается; '
                                            'попросить сократить текст')
        # роль для единого кегля по колоде
        role = sh.name if sh.name in ('Цифра', 'Подпись', 'Имя', 'Должность', 'Сноска', 'Подзаголовок', 'Модули',
                                      'Особенности', 'Описание', 'Подпись группы') else None
        if role is None and re.fullmatch(r'[\dхx][\dхx\s,.%+/×]*(из\s*[\dхx]+)?[\dхx%+]*', sh.text_frame.text.strip()) \
                and main >= 30:
            role = 'Цифра'
        if role:
            roles[role].append((no, max(sizes) if role == 'Цифра' else main))
        sh._rv_size = main
        # вылеты
        if widest > iw + 1:
            self.add(no, 'ОШИБКА', 'перенос', f'«{sh.text_frame.text[:30]}…»: слово шире рамки — разорвётся по буквам')
        if not sh.is_placeholder and need > ih + max(3, main * 0.3) and not xp(sh._element, './/a:bodyPr/a:spAutoFit'):
            self.add(no, 'ВНИМАНИЕ', 'вылет', f'«{sh.text_frame.text[:30]}…»: текст выше рамки на {need - ih:.0f} pt')

    def title_colour(self, sh):
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.text.strip():
                    try:
                        return str(r.font.color.rgb)
                    except Exception:
                        return 'по шаблону'
        return None

    def overlaps(self, no, slide, texts, icons, photos, decor=()):
        ext = [(sh, sh._rv_ext) for sh, _ in texts if hasattr(sh, '_rv_ext')]
        for i in range(len(ext)):
            for j in range(i + 1, len(ext)):
                w, h = inter(ext[i][1], ext[j][1])
                if w > 2 and h > 4:
                    a, b = ext[i][0].text_frame.text[:25], ext[j][0].text_frame.text[:25]
                    self.add(no, 'ОШИБКА', 'наложение', f'текст «{a}…» налезает на «{b}…» ({h:.0f} pt) — уменьшить кегль '
                                                         'по всей колоде или сократить')
        for sh, e in ext:
            for ic, box in icons:
                if transparent(ic):
                    continue
                w, h = inter(e, box)
                if w > 1 and h > 1:
                    self.add(no, 'ОШИБКА', 'наложение', f'текст «{sh.text_frame.text[:25]}…» налезает на иконку')
            for d, box in decor:  # водяной знак под текстом читается как наложение — так нельзя
                w, h = inter(e, box)
                if w > 1 and h > 1 and (stype(d) != 13 or covers(d, box, e, level=20)):
                    self.add(no, 'ОШИБКА', 'наложение', f'текст «{sh.text_frame.text[:25]}…» заходит на водяной знак '
                                                         f'«{d.name}» — сдвинуть или уменьшить знак')
            for ph, box in photos:
                near = (e[0] - 8, e[1], e[2] + 16, e[3])  # зазор до фото не меньше 8 pt
                w, h = inter(near, box)
                if w > 1 and h > 1 and not inside(e, box) and not transparent(ph) and covers(ph, box, near):
                    self.add(no, 'ОШИБКА', 'наложение', f'текст «{sh.text_frame.text[:25]}…» заходит на фото «{ph.name}» '
                                                         'или стоит впритык к нему')

    def picture_checks(self, no, photos, cards):
        for pic, box in photos:
            if not is_photo(pic):
                continue
            background = any(inside(getattr(t, '_rv_ext', b), box) for t, b in self._texts)
            d = 0 if background else aspect_error(pic, box)
            if d > 0.03:
                self.add(no, 'ОШИБКА', 'фото', f'«{pic.name}» растянуто на {d:.0%} — вернуть пропорции, кадрировать')
            for side, share, how in cut_edges(pic, box, self.sw, self.sh):
                if how == 'краем слайда':
                    self.add(no, 'ВНИМАНИЕ', 'фото', f'портрет «{pic.name}» уходит за край слайда {side} — '
                                                     'проверить глазами, что срез выглядит задуманным')
                else:
                    self.add(no, 'ОШИБКА', 'фото', f'портрет «{pic.name}» обрезан {how} {side} (≈{share:.0%} края) — '
                                                   'плечо или голова срезаны; уменьшить или сдвинуть, чтобы фигура '
                                                   'целиком помещалась (кроме низа)')
            for card, cb in cards:
                w, h = inter(box, cb)
                if w * h > 0.5 * box[2] * box[3] and not inside(box, cb, 1) and not inside(cb, box, 1):
                    self.add(no, 'ОШИБКА', 'фото', f'«{pic.name}» вылезает за край карточки')

    def icon_checks(self, no, icons, cards):
        info = []
        for ic, box in icons:
            d = group_distortion(ic) if stype(ic) == 6 else aspect_error(ic, box)
            if d > 0.04:
                self.add(no, 'ОШИБКА', 'иконка', f'иконка «{ic.name}» искажена: пропорции сбиты на {d:.0%}')
            card = min((c for c in cards if centre_in(box, c[1])), key=lambda c: c[1][2] * c[1][3], default=None)
            if card is None:
                continue
            cb = card[1]
            if not inside(box, cb, 0):
                self.add(no, 'ОШИБКА', 'иконка', f'иконка «{ic.name}» вылезает за карточку')
            elif min(box[0] - cb[0], box[1] - cb[1], cb[0] + cb[2] - box[0] - box[2]) < 6 and not transparent(ic):
                self.add(no, 'ВНИМАНИЕ', 'иконка', f'иконка «{ic.name}» прижата к краю карточки')
            info.append((ic, box, cb))
        # одинаковые карточки: иконки одного размера и на одном месте
        groups = collections.defaultdict(list)
        for ic, box, cb in info:
            if not transparent(ic):
                groups[(round(cb[2] / 4), round(cb[3] / 4))].append((ic, box, cb))
        for key, g in groups.items():
            if len(g) < 2:
                continue
            sizes = [max(b[2], b[3]) for _, b, _ in g]
            if max(sizes) - min(sizes) > 2:
                self.add(no, 'ОШИБКА', 'иконка', 'иконки в одинаковых карточках разного размера: '
                                                 + ', '.join(f'{s:.0f}' for s in sizes) + ' pt')
            offs = [(round(b[0] - c[0]), round(b[1] - c[1])) for _, b, c in g]
            common = collections.Counter(offs).most_common(1)[0][0]
            for (ic, b, c), off in zip(g, offs):
                if abs(off[0] - common[0]) > 2 or abs(off[1] - common[1]) > 2:
                    self.add(no, 'ОШИБКА', 'иконка', f'иконка «{ic.name}» сдвинута: отступ от края карточки {off}, '
                                                     f'у соседних {common}')
            same = [c for c in cards if (round(c[1][2] / 4), round(c[1][3] / 4)) == key]
            with_icon = {id(c) for _, _, c in g}
            empty = [c for c in same if not any(inside(b, c[1], 0) for _, b, _ in g)]
            if empty and len(empty) < len(same):
                for c, cb in empty:
                    self.add(no, 'ВНИМАНИЕ', 'иконка', f'в карточке ({cb[0]:.0f}, {cb[1]:.0f}) нет иконки, у соседних есть')

    def margins(self, no, sh, box):
        x, y, w, h = box
        if x <= 1 or y <= 1 or x + w >= self.sw - 1 or y + h >= self.sh - 1:
            return  # фон и картинки на вылет
        if x < LEFT - 1.5:
            self.add(no, 'ВНИМАНИЕ', 'поля', f'«{sh.name}» левее поля: x = {x:.0f} при {LEFT}')
        if x + w > RIGHT + 1.5:
            self.add(no, 'ВНИМАНИЕ', 'поля', f'«{sh.name}» правее поля: {x + w:.0f} при {RIGHT}')
        if y + h > GRID_BOTTOM + 2 and y < 480:
            self.add(no, 'ОШИБКА', 'поля', f'«{sh.name}» ниже сетки шаблона: {y + h:.0f} при {GRID_BOTTOM}')
        elif y + h > BOTTOM + 3 and y < 480 and self.house:
            self.add(no, 'ОШИБКА', 'поля', f'«{sh.name}» ниже нижней границы фирменного стиля: {y + h:.0f} при {BOTTOM}')

    def alignment(self, no, shapes, what):
        for i, edge in ((0, 'левые'), (1, 'верхние')):
            vals = sorted({round(b[i], 1) for _, b in shapes if 0 < b[i] < (self.sw if i == 0 else self.sh)})
            for a, b in zip(vals, vals[1:]):
                if 0.6 < b - a <= 3:
                    self.add(no, 'ВНИМАНИЕ', 'выравнивание', f'{edge} края {what} почти совпадают: {a:g} и {b:g} — выровнять')

    def deck_uniform(self, roles, titles):
        if titles:
            sizes = collections.defaultdict(list)
            for no, s, _ in titles:
                sizes[round(s, 1)].append(no)
            if len(sizes) > 1:
                self.add(0, 'ОШИБКА', 'кегль', 'заголовки разного кегля: ' + '; '.join(
                    f'{s:g} pt — слайды {", ".join(map(str, v))}' for s, v in sorted(sizes.items())) +
                    ' — привести к одному (если не влезает — уменьшить на всех)')
            cols = collections.defaultdict(list)
            for no, _, c in titles:
                cols[c].append(no)
            if len(cols) > 1:
                self.add(0, 'ОШИБКА', 'цвет', 'заголовки разного цвета: ' + '; '.join(
                    f'{c} — слайды {", ".join(map(str, v))}' for c, v in cols.items()) + ' — все цветом направления')
        for role, vals in roles.items():
            sizes = collections.defaultdict(set)
            for no, s in vals:
                sizes[round(s, 1)].add(no)
            if len(sizes) > 1:
                self.add(0, 'ОШИБКА', 'кегль', f'«{role}» разного кегля: ' + '; '.join(
                    f'{s:g} pt — слайды {", ".join(map(str, sorted(v)))}' for s, v in sorted(sizes.items())))
        # одинаковые поля в ряду (одного размера) — один кегль
        for no, slide in enumerate(self.prs.slides, 1):
            rows = collections.defaultdict(list)
            for sh, box in walk(slide.shapes):
                if sh.has_text_frame and sh.text_frame.text.strip() and hasattr(sh, '_rv_size') and not sh.is_placeholder:
                    rows[(round(box[2] / 3), round(box[3] / 3), round(box[1] / 3))].append(sh._rv_size)
            for k, v in rows.items():
                if len(v) > 1 and max(v) - min(v) > 0.4:
                    self.add(no, 'ОШИБКА', 'кегль', 'одинаковые поля в ряду разного кегля: '
                                                    + ', '.join(f'{s:g}' for s in v) + ' pt')

    # -- вывод
    def report(self, hanging=None, fixed=False):
        out = []
        errs = sum(1 for i in self.items if i['level'] == 'ОШИБКА')
        warns = len(self.items) - errs
        by = collections.defaultdict(list)
        seen = set()
        for i in self.items:
            key = (i['slide'], i['msg'])
            if key not in seen:
                seen.add(key)
                by[i['slide']].append(i)
        for no in sorted(by):
            out.append('колода:' if no == 0 else f'слайд {no}:')
            for i in sorted(by[no], key=lambda i: i['level'] != 'ОШИБКА'):
                out.append(f"  {i['level']} [{i['kind']}] {i['msg']}")
        if hanging:
            total = sum(hanging.values())
            verb = 'исправлено' if fixed else 'найдено (исправить: --fix)'
            out.append(f"висячие предлоги и союзы: {total} мест {verb} — слайды "
                       + ', '.join(f'{k} ({v})' for k, v in sorted(hanging.items()) if v))
        out.append(f'итог: ошибок {errs}, предупреждений {warns}' + (' — можно отдавать после визуальной проверки'
                                                                    if not errs else ' — исправить до отдачи'))
        return '\n'.join(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('pptx')
    ap.add_argument('--fix', metavar='OUT', help='сохранить копию с неразрывными пробелами после предлогов')
    ap.add_argument('--json', action='store_true')
    a = ap.parse_args()
    prs = Presentation(a.pptx)
    hanging = fix_hanging(prs) if a.fix else count_hanging(prs)
    if a.fix:
        prs.save(a.fix)
        prs = Presentation(a.fix)
    rv = Review(prs).run()
    if a.json:
        print(json.dumps(dict(items=rv.items, hanging=dict(hanging)), ensure_ascii=False, indent=1))
    else:
        print(rv.report(+hanging, bool(a.fix)))
    sys.exit(1 if any(i['level'] == 'ОШИБКА' for i in rv.items) else 0)


if __name__ == '__main__':
    main()
