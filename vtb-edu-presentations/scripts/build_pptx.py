#!/usr/bin/env python3
"""Редактируемая презентация PPTX по мастер-шаблону «ВТБ Образование» (мерико × ВТБ Образование, 2026).

    python3 build_pptx.py deck.json -o out/              # deck.pptx + report.json
    python3 build_pptx.py deck.json -o out/ --preview    # + deck.pdf и slide_XX.png (через LibreOffice)

Размер слайда 1920×1080 pt — как в шаблоне, поэтому кегли совпадают с гайдом один в один:
заголовок слайда 82, титульный 128, большие цифры 120, текст 16–42, минимум 16.
Формат deck.json и макеты — references/layouts.md, правила бренда — references/brand.md.
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageFont
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt
from lxml import etree

SKILL = Path(__file__).resolve().parent.parent
A = SKILL / "assets"
W, H = 1920, 1080
PX = 12700  # 1 px макета = 1 pt = 12700 EMU

# --- Шрифты ---------------------------------------------------------------
# Фирменный шрифт — VTB Group (начертания Book и Light). Имена семейств можно
# переопределить флагами --font / --font-light, если в системе они называются иначе.
FONT = "VTB Group Book"
FONT_LIGHT = "VTB Group Light"
# Метрики для оценки переноса строк: Open Sans близок по ширине знаков к VTB Group.
MEASURE = {False: A / "fonts" / "OpenSans-Regular.ttf", True: A / "fonts" / "OpenSans-Light.ttf"}

# --- Палитра (references/brand.md) ---------------------------------------
RETAIL = "0A2896"     # Retail Blue
MEDIUM = "1E4BD2"     # Medium Blue
CYAN = "00AAFF"       # Cyan
RETAIL20 = "CCD4E6"   # Retail Blue 20 %
BLACK = "000000"
WHITE = "FFFFFF"
GREY = "ACB6C3"       # вторая часть заголовка, подписи
TEXT = "6B7683"       # основной текст на светлом фоне
CARD = "F4F6FA"       # тело карточки (EAEDF5 на 50 %)
CARD_HEAD = "EAEDF5"  # шапка серой карточки, выделенная строка таблицы
PAGE = "C9CFD9"       # номер слайда

# Направления: градиент шапки (слева → направо), второй цвет, сплошной цвет, акцент цифр, ряд для графиков.
DIRECTIONS = {
    "vtb":       dict(name="ВТБ Образование", grad=("1E4BD2", "86BFF5"), second=CYAN, solid=MEDIUM, accent=MEDIUM,
                      chart=[RETAIL, MEDIUM, CYAN, GREY]),
    "family":    dict(name="Финансовая грамотность для всей семьи", grad=("00AAFF", "803CE9"), second="803CE9",
                      solid=MEDIUM, accent=MEDIUM, chart=[MEDIUM, "803CE9", "F7A2F7", CYAN]),
    "prof":      dict(name="Профессиональные финансы", grad=(RETAIL, "750CF6"), second="8F95FF", solid=RETAIL,
                      accent="6218D3", chart=[RETAIL, "750CF6", "8F95FF", GREY]),
    "china":     dict(name="Китайское направление", grad=(RETAIL, "D8A2F7"), second="F5C482", solid=MEDIUM,
                      accent=MEDIUM, chart=[RETAIL, MEDIUM, "D8A2F7", "F5C482"]),
    "b2b":       dict(name="B2B-направление", grad=(MEDIUM, "FFB68E"), second="FFB68E", solid=MEDIUM, accent=MEDIUM,
                      chart=[MEDIUM, "FFB68E", RETAIL20, RETAIL]),
    "infra":     dict(name="Крупные инфраструктурные проекты в регионах РФ", grad=(MEDIUM, CYAN), second=CYAN,
                      solid=MEDIUM, accent=MEDIUM, chart=[MEDIUM, CYAN, RETAIL20, RETAIL]),
    "private":   dict(name="Private Banking", grad=("000324", "57637A"), second="D3875F", solid="000324",
                      accent="57637A", chart=["000324", "57637A", GREY, "D3875F"]),
    "privilege": dict(name="Привилегия", grad=("002043", "900258"), second="900258", solid="002043",
                      accent="900258", chart=["002043", "900258", "48114D", GREY]),
}

T = DIRECTIONS["vtb"]  # активное направление, задаётся в main()
REPORT = []            # предупреждения о тексте, который пришлось уменьшить или который не влез


# --- Низкоуровневые помощники ---------------------------------------------
def E(v):
    return Emu(int(round(v * PX)))


def rgb(h):
    return RGBColor.from_string(h)


def _alpha(color_el, opacity):
    if opacity is None or opacity >= 1:
        return
    a = etree.SubElement(color_el, qn("a:alpha"))
    a.set("val", str(int(opacity * 100000)))


def fill_solid(shape, color, opacity=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)
    if opacity is not None:
        _alpha(shape.fill._xPr.find(qn("a:solidFill"))[0], opacity)


def fill_grad(shape, c1, c2, angle=0, mid=None):
    """Линейный градиент. angle: 0 — слева направо, 90 — сверху вниз."""
    f = shape.fill
    f.gradient()
    f.gradient_angle = angle
    st = f.gradient_stops
    st[0].color.rgb, st[0].position = rgb(c1), 0
    st[1].color.rgb, st[1].position = rgb(c2), 1.0
    if mid:  # промежуточная точка: цвет держится дольше у начала, как в шаблоне
        gs_lst = f._xPr.find(qn("a:gradFill")).find(qn("a:gsLst"))
        gs = etree.SubElement(gs_lst, qn("a:gs"))
        gs.set("pos", str(int(mid[1] * 100000)))
        clr = etree.SubElement(gs, qn("a:srgbClr"))
        clr.set("val", mid[0])
        gs_lst[:] = sorted(gs_lst, key=lambda g: int(g.get("pos")))


def no_line(shape):
    shape.line.fill.background()


def plain_style(shape):
    """Убрать ссылку на стиль темы: иначе часть программ дорисовывает тень и обводку."""
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def rect(slide, x, y, w, h, color=None, grad=None, radius=0, opacity=None, line=None, angle=0, top_only=False):
    kind = (MSO_SHAPE.ROUND_2_SAME_RECTANGLE if top_only else MSO_SHAPE.ROUNDED_RECTANGLE) if radius \
        else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(kind, E(x), E(y), E(w), E(h))
    if radius:
        s.adjustments[0] = min(0.5, radius / min(w, h))
        if top_only:
            s.adjustments[1] = 0
    if grad:
        fill_grad(s, *grad[:2], angle=angle, mid=grad[2] if len(grad) > 2 else None)
    elif color:
        fill_solid(s, color, opacity)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line[0])
        s.line.width = E(line[1])
    else:
        no_line(s)
    plain_style(s)
    return s


def line(slide, x1, y1, x2, y2, color=WHITE, width=1.5, opacity=None):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2))
    c.line.color.rgb = rgb(color)
    c.line.width = E(width)
    if opacity is not None:
        _alpha(c.line._ln.find(qn("a:solidFill"))[0], opacity)
    plain_style(c)
    return c


def polygon(slide, pts, color=WHITE):
    fb = slide.shapes.build_freeform(E(pts[0][0]), E(pts[0][1]), scale=1.0)
    fb.add_line_segments([(E(x), E(y)) for x, y in pts[1:]], close=True)
    s = fb.convert_to_shape()
    fill_solid(s, color)
    no_line(s)
    plain_style(s)
    return s


def picture(slide, path, x, y, w=None, h=None, cover=False, radius=0):
    """cover=True — заполнить рамку w×h с обрезкой, как object-fit: cover."""
    path = str(path)
    if cover and w and h:
        iw, ih = Image.open(path).size
        pic = slide.shapes.add_picture(path, E(x), E(y), E(w), E(h))
        r_img, r_box = iw / ih, w / h
        if r_img > r_box:
            c = (1 - r_box / r_img) / 2
            pic.crop_left = pic.crop_right = c
        else:
            c = (1 - r_img / r_box) / 2
            pic.crop_top = pic.crop_bottom = c
    else:
        pic = slide.shapes.add_picture(path, E(x), E(y), E(w) if w else None, E(h) if h else None)
    if radius:
        pic.auto_shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
        geom = pic._element.spPr.find(qn("a:prstGeom"))
        av = geom.find(qn("a:avLst"))
        if av is None:
            av = etree.SubElement(geom, qn("a:avLst"))
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", "adj")
        gd.set("fmla", f"val {int(radius / min(pic.width, pic.height) * PX * 100000)}")
    return pic


# --- Текст ----------------------------------------------------------------
_fonts = {}


def _font(size, light):
    k = (round(size * 4) / 4, light)
    if k not in _fonts:
        _fonts[k] = ImageFont.truetype(str(MEASURE[light]), int(round(size * 4)))
    return _fonts[k]


def measure(text, size, width, light=False, upper=False, spacing=1.2):
    """Сколько строк и какой высоты займёт текст в рамке шириной width (px)."""
    f = _font(size, light)
    lines = 0
    for para in re.split(r"<br\s*/?>|\n", str(text)):
        para = typo(plain(para))  # те же неразрывные пробелы, что попадут в слайд
        if upper:
            para = para.upper()
        words = para.split(" ")
        cur = ""
        n = 1
        for wd in words:
            if f.getlength(wd) / 4 > width:  # слово (или группа через неразрывный пробел) шире рамки — разорвётся
                return 99, 99 * size * spacing
            t = (cur + " " + wd).strip()
            if f.getlength(t) / 4 <= width:
                cur = t
            else:
                n += 1
                cur = wd
        lines += n
    return lines, lines * size * spacing


def plain(s):
    s = re.sub(r"<br\s*/?>", "\n", str(s))
    s = re.sub(r"<[^>]+>", "", s)
    return s.replace("&nbsp;", " ").replace("&amp;", "&")


_SHORT = re.compile(r"(?<![\w])([A-Za-zА-Яа-яЁё]{1,2}|—|–)\s+(?=\S)")


def typo(s):
    """Неразрывный пробел после коротких слов и перед тире — как в материалах ВТБ Образования."""
    s = _SHORT.sub(lambda m: m.group(1) + " ", s)
    return re.sub(r"\s+—", " —", s)


def runs(text):
    """Разбор мини-разметки: <b>…</b> — акцентный цвет, <g>…</g> — серый, \\n или <br> — новая строка."""
    out = []
    for para in re.split(r"<br\s*/?>|\n", str(text)):
        parts = []
        for m in re.finditer(r"<(b|g)>(.*?)</\1>|([^<]+)", para):
            if m.group(3) is not None:
                parts.append((typo(plain(m.group(3))), None))
            else:
                parts.append((typo(plain(m.group(2))), m.group(1)))
        out.append(parts)
    return out


def text(slide, x, y, w, h, s, size=32, color=TEXT, light=False, upper=False, align="l", anchor="t",
         spacing=1.2, min_size=16, accent=None, where=""):
    """Текстовый блок. Если текст не влезает — уменьшает кегль до min_size и пишет в отчёт."""
    s = "" if s is None else str(s)
    if upper:
        s = re.sub(r"(<[^>]+>)|([^<]+)", lambda m: m.group(1) or m.group(2).upper(), s)
    sz = size
    while sz > min_size and measure(s, sz, w, light, spacing=spacing)[1] > h + 1:
        sz -= 1
    if sz != size:
        REPORT.append(f"{where}: кегль уменьшен {size}→{sz}")
    if measure(s, sz, w, light, spacing=spacing)[1] > h + 1:
        REPORT.append(f"{where}: ПЕРЕПОЛНЕНИЕ — текст не влезает даже при {sz} pt, сократите")
    tb = slide.shapes.add_textbox(E(x), E(y), E(w), E(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    for i, parts in enumerate(runs(s)):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = spacing
        for t, kind in parts or [("", None)]:
            r = p.add_run()
            r.text = t
            f = r.font
            f.size = Pt(sz)
            f.name = FONT_LIGHT if light else FONT
            f.color.rgb = rgb({"b": accent or T["accent"], "g": GREY}.get(kind, color))
    return tb


def bullets(slide, x, y, w, h, items, size=24, color=TEXT, marker="•", where="", numbered=False):
    """Список: маркер «•» или номера 01, 02 акцентным цветом, как в карточках шаблона."""
    lines = []
    for i, it in enumerate(items, 1):
        m = f"{i:02d}" if numbered else marker
        lines.append(f"<b>{m}</b> {it}")
    return text(slide, x, y, w, h, "\n".join(lines), size=size, color=color, where=where, spacing=1.2)


# --- Каркас слайда --------------------------------------------------------
def asset(kind, name):
    """Путь к фирменному ресурсу или к файлу пользователя (относительно deck.json)."""
    if not name:
        return None
    p = Path(name)
    if p.suffix and (CTX["base"] / p).exists():
        return CTX["base"] / p
    if p.is_absolute() and p.exists():
        return p
    folder = {"portrait": "portraits", "photo": "photos", "icon": "icons", "bg": "backgrounds"}[kind]
    for ext in (".png", ".jpg"):
        q = A / folder / f"{name}{ext}"
        if q.exists():
            return q
    sys.exit(f"Не найден ресурс {kind}: {name}. Список — references/assets.md")


CTX = {}


def logo(slide, x, y, variant="color", h=50):
    p = A / "logo" / f"logo-{variant}.png"
    iw, ih = Image.open(p).size
    w = h * iw / ih
    slide.shapes.add_picture(str(p), E(x), E(y), E(w), E(h))
    return w


def logo_right(slide, variant="color"):
    p = A / "logo" / f"logo-{variant}.png"
    iw, ih = Image.open(p).size
    h = 50
    logo(slide, 1860 - h * iw / ih, 60, variant, h)


def partner(slide, s, x, y, h=50, color=BLACK):
    """Партнёрский знак через «×»: расстояние до «×» — половина высоты буквы «Т» логотипа."""
    p = s.get("partner_logo")
    if not p:
        return
    p = asset("photo", p)
    text(slide, x, y + h * 0.18, 30, h * 0.64, "×", size=30, color=color, align="c", anchor="m", min_size=30)
    iw, ih = Image.open(p).size
    slide.shapes.add_picture(str(p), E(x + 52), E(y), E(h * iw / ih), E(h))


def page_no(slide, n):
    text(slide, 1780, 1022, 80, 30, f"{n:02d}", size=16, color=PAGE, align="r", min_size=16)


def title(slide, s, where, top=52, width=1480, size=82):
    """Заголовок слайда: прописные, 82 pt, чёрный; часть в <g>…</g> — серая (как «РЕЗЮМЕ <g>ВТБ ОБРАЗОВАНИЕ</g>»).
    Возвращает Y начала контента (217 при одной строке)."""
    t = s.get("title", "")
    if s.get("title_grey"):
        t = f"{t} <g>{s['title_grey']}</g>"
    n, hgt = measure(t, size, width, upper=True, spacing=1.0)
    if n > 2:
        REPORT.append(f"{where}: заголовок в {n} строки — сократите до 1–2")
    text(slide, 60, top, width, max(hgt, size * 1.25), t, size=size, color=BLACK, upper=True, spacing=1.0,
         min_size=60, where=where + " заголовок")
    return 217 if n <= 1 else 217 + int(size * (min(n, 2) - 1))


def light_slide(prs, s, where, n):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    logo_right(sl)
    top = title(sl, s, where) if s.get("title") else 217
    page_no(sl, n)
    return sl, top


# --- Макеты: титул и разделы ------------------------------------------------
def guides(slide, ys=(380, 700), x=66, color=WHITE):
    """Тонкие линии сетки — фирменный приём титульника («структура знаний»)."""
    line(slide, x, 150, x, 1080, color, 1.5, 0.45)
    for y in ys:
        line(slide, 0, y, 1130, y, color, 1.5, 0.35)


def L_title(prs, s, n, where):
    """Титульный: фон-паттерн направления, белый логотип, портрет справа (variant: photo | plain | split)."""
    v = s.get("variant", "photo" if s.get("portrait") else "plain")
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg = asset("bg", s.get("background") or CTX["direction"])
    if v == "split":
        picture(sl, bg, 960, 0, 960, 1080, cover=True)
        if s.get("portrait"):
            picture(sl, asset("portrait", s["portrait"]), 960, 40, 960, 1040, cover=True)
        w = logo(sl, 60, 60, "color")
        partner(sl, s, 60 + w + 26, 60)
        text(sl, 60, 300, 860, 480, s.get("title", ""), size=82, color=BLACK, upper=True, spacing=1.0,
             anchor="m", min_size=56, where=where)
        bottom_caption(sl, s, where)
        return sl
    picture(sl, bg, 0, 0, W, H, cover=True)
    logo(sl, 60, 60, "white")
    line(sl, 400, 85, 1630, 85, WHITE, 1.5, 0.5)
    if s.get("meta"):
        text(sl, 1460, 52, 400, 70, s["meta"], size=24, color=WHITE, align="r", spacing=1.2, min_size=18, where=where)
    tw = 1080 if s.get("portrait") else 1740
    if s.get("portrait"):
        p = asset("portrait", s["portrait"])
        iw, ih = Image.open(p).size
        ph = 920
        pw = ph * iw / ih
        picture(sl, p, 1920 - pw + min(pw * 0.08, 80), 1080 - ph, pw, ph)
    guides(sl)
    text(sl, 62, 380, tw, 320, s.get("title", ""), size=128, color=WHITE, upper=True, spacing=0.92,
         anchor="b", min_size=72, where=where + " заголовок")
    if s.get("subtitle"):
        text(sl, 62, 775, tw, 110, s["subtitle"], size=42, color=WHITE, spacing=1.2, min_size=28, where=where)
    return sl


def bottom_caption(slide, s, where):
    """Подпись внизу слева: синяя вертикальная черта + название программы и дата."""
    lines = [x for x in (s.get("subtitle"), s.get("date")) if x]
    if not lines:
        return
    rect(slide, 60, 935, 4, 85, MEDIUM)
    t = s.get("subtitle", "")
    if s.get("date"):
        t = (t + "\n" if t else "") + f"<g>{s['date']}</g>"
    text(slide, 84, 942, 800, 75, t, size=24, color=RETAIL, spacing=1.2, min_size=18, where=where)


def L_section(prs, s, n, where):
    """Разделительный: белое поле слева, справа диагональный срез с паттерном или фото."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    img = asset("photo", s["photo"]) if s.get("photo") else asset("bg", s.get("background") or CTX["direction"])
    picture(sl, img, 860, 0, 1060, 1080, cover=True)
    polygon(sl, [(850, 0), (1300, 0), (858, 1080), (850, 1080)], WHITE)
    w = logo(sl, 60, 60, "color")
    partner(sl, s, 60 + w + 26, 60)
    num = f"<b>{s['num']}</b>\n" if s.get("num") else ""
    text(sl, 60, 300, 900, 480, num + s.get("title", ""), size=96, color=BLACK, upper=True, spacing=1.0,
         anchor="m", min_size=60, where=where)
    bottom_caption(sl, s, where)
    page_no(sl, n)
    return sl


def L_toc(prs, s, n, where):
    """Содержание: сетка 3×3 — крупная цифра, название раздела, страница."""
    sl, top = light_slide(prs, s, where, n)
    items = s["items"]
    cols = 3
    rows = (len(items) + cols - 1) // cols
    gw, gh = 20, 20
    cw = (1800 - gw * (cols - 1)) / cols
    ch = min(184, (1000 - top - gh * (rows - 1)) / rows)
    y0 = top + max(0, (1000 - top) - rows * ch - (rows - 1) * gh) / 2 if s.get("center") else top + 120
    for i, it in enumerate(items):
        it = it if isinstance(it, dict) else {"title": it}
        x = 60 + (i % cols) * (cw + gw)
        y = y0 + (i // cols) * (ch + gh)
        rect(sl, x, y, cw, ch, CARD, radius=12)
        text(sl, x + 30, y, 120, ch, str(it.get("num", i + 1)), size=147, color=T["accent"], light=True,
             anchor="m", spacing=1.0, min_size=100, where=where)
        text(sl, x + 150, y + 30, cw - 180, ch - 80, it["title"], size=32, color=BLACK, min_size=22, where=where)
        if it.get("page"):
            text(sl, x + 150, y + ch - 52, cw - 180, 36, f"Стр. {it['page']}", size=28, color=GREY, min_size=20)
    return sl


# --- Макеты: контент --------------------------------------------------------
def icon_pic(slide, name, x, y, size=64, white=False):
    p = asset("icon", name)
    if white:
        wp = p.with_name(p.stem + "-white.png")
        p = wp if wp.exists() else p
    slide.shapes.add_picture(str(p), E(x), E(y), E(size), E(size))


def L_bullets(prs, s, n, where):
    """Буллиты: строки-плашки с номером 01…, иконкой или без; conclusion — итог на голубой плашке."""
    sl, top = light_slide(prs, s, where, n)
    items = list(s["items"])
    concl = s.get("conclusion")
    rows = len(items) + (1 if concl else 0)
    g = 20
    h = (1000 - top - g * (rows - 1)) / rows
    for i, it in enumerate(items):
        it = it if isinstance(it, dict) else {"text": it}
        y = top + i * (h + g)
        rect(sl, 60, y, 1800, h, CARD, radius=12)
        if it.get("icon"):
            icon_pic(sl, it["icon"], 118, y + h / 2 - 32)
            tx = 240
        elif s.get("numbered", True):
            text(sl, 100, y, 100, h, f"{i + 1:02d}", size=62, color=T["accent"], light=True, anchor="m",
                 spacing=1.0, min_size=40)
            tx = 208
        else:
            tx = 100
        text(sl, tx, y + 24, 1560 - tx + 60, h - 48, it["text"], size=32, color=TEXT, anchor="m",
             min_size=20, where=f"{where} пункт {i + 1}")
    if concl:
        y = top + len(items) * (h + g)
        rect(sl, 60, y, 1800, h, T["second"], radius=12)
        text(sl, 100, y + 24, 1720, h - 48, concl, size=32, color=WHITE, anchor="m", min_size=20,
             where=where + " вывод")
    return sl


def header_style(i, n, hl=None):
    """Шапки карточек в порядке шаблона: серая → второй цвет → градиент → сплошной."""
    if hl is not None:
        return "grad" if i == hl else "grey"
    seq = {1: ["grad"], 2: ["grey", "grad"], 3: ["grey", "second", "grad"]}.get(n, ["grey", "second", "grad", "solid"])
    return seq[i % len(seq)]


def card_head(sl, x, y, w, h, style, heading, icon=None, size=32, where=""):
    kw = dict(radius=12, top_only=True)
    if style == "grey":
        rect(sl, x, y, w, h, CARD_HEAD, **kw)
        col = TEXT
    elif style == "second":
        rect(sl, x, y, w, h, T["second"], **kw)
        col = WHITE
    elif style == "solid":
        rect(sl, x, y, w, h, T["solid"], **kw)
        col = WHITE
    else:
        rect(sl, x, y, w, h, grad=(T["grad"][0], T["grad"][1], (T["grad"][0], 0.25)), **kw)
        col = WHITE
    tx = x + 30
    if icon:
        icon_pic(sl, icon, x + 28, y + h / 2 - 32, 64, white=(style != "grey"))
        tx = x + 118
    text(sl, tx, y + 14, x + w - tx - 24, h - 28, heading, size=size, color=col, anchor="m", min_size=20,
         where=where)


def L_cards(prs, s, n, where):
    """2–4 вертикальные карточки: шапка (иконка + заголовок), текст или список, фото снизу по желанию."""
    sl, top = light_slide(prs, s, where, n)
    items = s["items"]
    k = len(items)
    g = 20
    cw = (1800 - g * (k - 1)) / k
    head_h = 124 if k <= 3 else 110
    photo_h = 300 if any(it.get("photo") for it in items) else 0
    hsz = 32 if k <= 2 else 28
    for i, it in enumerate(items):
        x = 60 + i * (cw + g)
        style = header_style(i, k, s.get("highlight"))
        body_bottom = 1000 - (photo_h + g if photo_h else 0)
        rect(sl, x, top, cw, body_bottom - top, CARD, radius=12)
        card_head(sl, x, top, cw, head_h, style, it.get("title", ""), it.get("icon"), size=hsz,
                  where=f"{where} карточка {i + 1}")
        by = top + head_h + 30
        bh = body_bottom - by - 30
        if it.get("items"):
            bullets(sl, x + 30, by, cw - 60, bh, it["items"], size=24, numbered=it.get("numbered", False),
                    where=f"{where} карточка {i + 1}")
        else:
            text(sl, x + 30, by, cw - 60, bh, it.get("text", ""), size=24, color=TEXT, min_size=16,
                 where=f"{where} карточка {i + 1}")
        if photo_h:
            if it.get("photo"):
                picture(sl, asset("photo", it["photo"]), x, 1000 - photo_h, cw, photo_h, cover=True, radius=12)
    return sl


def L_process(prs, s, n, where):
    """Карточка с процессом: шаги «Шаг 1…», последний — «Результат» на градиенте."""
    sl, top = light_slide(prs, s, where, n)
    items = s["items"]
    k = len(items)
    g = 20
    cw = (1800 - g * (k - 1)) / k
    for i, it in enumerate(items):
        x = 60 + i * (cw + g)
        last = i == k - 1 and s.get("result", True)
        rect(sl, x, top, cw, 1000 - top, CARD, radius=12)
        lab = it.get("label") or ("Результат" if last else f"Шаг {i + 1}")
        if last:
            rect(sl, x, top, cw, 74, grad=(T["grad"][0], T["grad"][1], (T["grad"][0], 0.3)), radius=12, top_only=True)
        else:
            rect(sl, x, top, cw, 74, CARD_HEAD, radius=12, top_only=True)
        text(sl, x + 30, top, cw - 60, 74, lab, size=32, color=WHITE if last else TEXT, anchor="m", min_size=22)
        text(sl, x + 30, top + 104, cw - 60, 80, it.get("title", ""), size=28,
             color=T["accent"] if last else BLACK, min_size=20, where=f"{where} шаг {i + 1}")
        text(sl, x + 30, top + 200, cw - 60, 1000 - top - 230, it.get("text", ""), size=24, color=TEXT,
             min_size=16, where=f"{where} шаг {i + 1}")
    return sl


def L_highlights(prs, s, n, where):
    """Хайлайты: 4–8 ячеек с крупной цифрой (120 pt) или иконкой и коротким текстом."""
    sl, top = light_slide(prs, s, where, n)
    items = s["items"]
    k = len(items)
    cols = s.get("columns") or (3 if k in (3, 5, 6, 9) else 4 if k in (4, 7, 8) else 2)
    rows = (k + cols - 1) // cols
    g = 20
    cw = (1800 - g * (cols - 1)) / cols
    ch = (1000 - top - g * (rows - 1)) / rows
    for i, it in enumerate(items):
        x = 60 + (i % cols) * (cw + g)
        y = top + (i // cols) * (ch + g)
        if it.get("highlight"):
            rect(sl, x, y, cw, ch, grad=("9DDDFF", "4F92E8"), radius=12, angle=90)
            col, tcol = WHITE, WHITE
        else:
            rect(sl, x, y, cw, ch, CARD, radius=12)
            col, tcol = T["accent"], TEXT
        if it.get("value") is not None:
            unit = f"<g>{it['unit']}</g>" if it.get("unit") else ""
            text(sl, x + 40, y + 20, cw - 80, 150, f"{it['value']}{unit}", size=120 if cols <= 3 else 100,
                 color=col, light=True, spacing=1.0, min_size=56, where=f"{where} ячейка {i + 1}")
        elif it.get("icon"):
            icon_pic(sl, it["icon"], x + 40, y + 40, 80, white=bool(it.get("highlight")))
        text(sl, x + 40, y + ch * 0.55, cw - 80, ch * 0.45 - 30, it.get("text", ""), size=24, color=tcol,
             anchor="b", min_size=16, where=f"{where} ячейка {i + 1}")
    return sl


def L_photo_text(prs, s, n, where):
    """Слайд с фото и текстом: карточка со шапкой и списком с иконками + фото (photo_side: right | left)."""
    sl, top = light_slide(prs, s, where, n)
    left = s.get("photo_side", "right") == "left"
    cx, px = (970, 60) if left else (60, 970)
    rect(sl, cx, top, 890, 1000 - top, CARD, radius=12)
    hh = 70
    rect(sl, cx, top, 890, hh, grad=(T["grad"][0], T["grad"][1], (T["grad"][0], 0.35)), radius=12, top_only=True)
    text(sl, cx + 30, top, 830, hh, s.get("label", ""), size=28, color=WHITE, upper=True, anchor="m", min_size=20)
    items = s.get("items", [])
    k = max(1, len(items))
    y0 = top + hh + 30
    rh = (1000 - y0 - 20) / k
    for i, it in enumerate(items):
        it = it if isinstance(it, dict) else {"text": it}
        y = y0 + i * rh
        tx = cx + 30
        if it.get("icon"):
            icon_pic(sl, it["icon"], cx + 30, y + 6, 48)
            tx = cx + 108
        text(sl, tx, y, cx + 860 - tx, rh - 12, it["text"], size=26, color=TEXT, min_size=16,
             where=f"{where} пункт {i + 1}")
    picture(sl, asset("photo", s["photo"]), px, top, 890, 1000 - top, cover=True, radius=12)
    return sl


def L_menu_photo(prs, s, n, where):
    """4 карточки + фото: слева пункты-меню (иконка, заголовок, подпись), последний выделен; справа фото."""
    sl, top = light_slide(prs, s, where, n)
    items = s["items"]
    k = len(items)
    g = 20
    h = (1000 - top - g * (k - 1)) / k
    for i, it in enumerate(items):
        y = top + i * (h + g)
        hl = it.get("highlight", i == k - 1)
        if hl:
            rect(sl, 60, y, 690, h, grad=(T["grad"][0], T["grad"][1], (T["grad"][0], 0.3)), radius=12)
        else:
            rect(sl, 60, y, 690, h, CARD_HEAD, radius=12)
        if it.get("icon"):
            icon_pic(sl, it["icon"], 90, y + h / 2 - 32, 64, white=hl)
        text(sl, 180, y + 20, 540, h * 0.5 - 10, it.get("title", ""), size=38, color=WHITE if hl else T["accent"],
             anchor="b", min_size=24, where=f"{where} пункт {i + 1}")
        text(sl, 180, y + h * 0.5 + 8, 540, h * 0.5 - 28, it.get("text", ""), size=24,
             color=WHITE if hl else TEXT, min_size=16, where=f"{where} пункт {i + 1}")
    picture(sl, asset("photo", s["photo"]), 770, top, 1090, 1000 - top, cover=True, radius=12)
    return sl


def L_table(prs, s, n, where):
    """Таблица: шапка Retail Blue / Medium Blue / Cyan / серый, зебра, выделенные строки — синим текстом.
    side — карточка справа с цифрой и текстом, note — плашка-вывод под таблицей."""
    sl, top = light_slide(prs, s, where, n)
    headers, rows_ = s["headers"], s["rows"]
    side, note = s.get("side"), s.get("note")
    tw = 1435 if side else 1800
    nh = 120 if note else 0
    avail = 1000 - top - (nh + 20 if note else 0)
    nr = len(rows_) + 1
    rh = min(74, avail / nr)
    fs = 28 if rh >= 64 else 24 if rh >= 52 else 20
    shape = sl.shapes.add_table(nr, len(headers), E(60), E(top), E(tw), E(rh * nr))
    tbl = shape.table
    tbl.first_row = True
    tblPr = shape._element.graphic.graphicData.tbl.tblPr
    st = tblPr.find(qn("a:tableStyleId"))
    if st is not None:
        st.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # «Без стиля, сетка» — цвета задаём сами
    widths = s.get("widths")
    for j in range(len(headers)):
        tbl.columns[j].width = E((widths[j] if widths else (tw * (0.3 if j == 0 else 0.7 / (len(headers) - 1)))))
    for r in range(nr):
        tbl.rows[r].height = E(rh)
    hcol = [RETAIL, MEDIUM, CYAN, GREY]
    align = s.get("align", [])
    hl_rows = set(s.get("highlight_rows", []))
    for r in range(nr):
        for j in range(len(headers)):
            c = tbl.cell(r, j)
            c.margin_left = c.margin_right = E(20)
            c.margin_top = c.margin_bottom = E(4)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            if r == 0:
                c.fill.solid()
                c.fill.fore_color.rgb = rgb(hcol[0] if j == 0 else hcol[1 + (j - 1) % 3])
                val, col = headers[j], WHITE
            else:
                hl = (r - 1) in hl_rows
                c.fill.solid()
                c.fill.fore_color.rgb = rgb(CARD_HEAD if hl else (CARD if r % 2 else WHITE))
                val, col = rows_[r - 1][j], (T["accent"] if hl else TEXT)
            no_borders(c)
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            a = align[j] if j < len(align) else ("l" if j == 0 else "c")
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[a[0]]
            run = p.add_run()
            run.text = typo(plain(val))
            run.font.size = Pt(fs)
            run.font.name = FONT
            run.font.color.rgb = rgb(col)
            if measure(run.text, fs, tbl.columns[j].width / PX - 40)[0] > max(1, int(rh // (fs * 1.25))):
                REPORT.append(f"{where}: ячейка [{r},{j}] переносится — строка станет выше, проверьте высоту таблицы")
    if side:
        x = 1515
        rect(sl, x, top, 345, 1000 - top - (nh + 20 if note else 0), grad=("9DDDFF", "1E4BD2"), radius=12, angle=90)
        text(sl, x + 30, top + 30, 285, 110, side.get("value", ""), size=82, color=WHITE, light=True, spacing=1.0,
             min_size=48)
        text(sl, x + 30, top + 160, 285, 1000 - top - 200 - (nh + 20 if note else 0), side.get("text", ""),
             size=24, color=WHITE, min_size=16, where=where + " боковая карточка")
    if note:
        y = 1000 - nh
        rect(sl, 60, y, 1000, nh, grad=(T["grad"][0], T["grad"][1], (T["grad"][0], 0.35)), radius=12)
        text(sl, 90, y + 14, 940, nh - 28, note, size=24, color=WHITE, anchor="m", min_size=16,
             where=where + " вывод")
    if s.get("source"):
        text(sl, 60, 1010, 1500, 40, f"Источник: {s['source']}", size=16, color=GREY, min_size=16)
    return sl


def no_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcPr, qn(tag))
        ln.set("w", "0")
        etree.SubElement(ln, qn("a:noFill"))
    # порядок: линии должны идти до заливки
    fill = [e for e in tcPr if e.tag in (qn("a:solidFill"), qn("a:noFill")) and e.getparent() is tcPr]
    for f in fill:
        tcPr.remove(f)
        tcPr.append(f)


CHARTS = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "stacked": XL_CHART_TYPE.COLUMN_STACKED,
          "hbar": XL_CHART_TYPE.BAR_CLUSTERED, "line": XL_CHART_TYPE.LINE_MARKERS,
          "donut": XL_CHART_TYPE.DOUGHNUT, "pie": XL_CHART_TYPE.PIE}


def chart(sl, x, y, w, h, c):
    cd = CategoryChartData()
    cd.categories = c["categories"]
    for se in c["series"]:
        cd.add_series(se["name"], se["values"])
    kind = c.get("type", "bar")
    gf = sl.shapes.add_chart(CHARTS[kind], E(x), E(y), E(w), E(h), cd)
    ch = gf.chart
    ch.has_title = False
    ch.font.size = Pt(18)
    ch.font.name = FONT
    ch.font.color.rgb = rgb(TEXT)
    pal = T["chart"]
    multi = len(c["series"]) > 1 or kind in ("donut", "pie")
    ch.has_legend = multi
    if multi:
        ch.legend.position = XL_LEGEND_POSITION.TOP if kind not in ("donut", "pie") else XL_LEGEND_POSITION.RIGHT
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(18)
    plot = ch.plots[0]
    plot.has_data_labels = c.get("labels", True)
    if plot.has_data_labels:
        dl = plot.data_labels
        dl.font.size = Pt(18)
        dl.number_format = c.get("format", "General")
        dl.number_format_is_linked = False
        if kind in ("bar", "hbar"):
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
    if kind in ("donut", "pie"):
        for i, pt in enumerate(plot.series[0].points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb(pal[i % len(pal)])
    else:
        if kind in ("bar", "stacked", "hbar"):
            plot.gap_width = 60
            if kind == "stacked":
                plot.overlap = 100
        for i, se in enumerate(plot.series):
            col = pal[i % len(pal)]
            if kind == "line":
                se.format.line.color.rgb = rgb(col)
                se.format.line.width = Pt(4)
                se.smooth = False
            else:
                se.format.fill.solid()
                se.format.fill.fore_color.rgb = rgb(col)
        va = ch.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = rgb("E3E7EF")
        va.format.line.fill.background()
        va.tick_labels.font.size = Pt(16)
        va.tick_labels.font.color.rgb = rgb(GREY)
        if c.get("max"):
            va.maximum_scale = c["max"]
        if c.get("min") is not None:
            va.minimum_scale = c["min"]
        if c.get("step"):
            va.major_unit = c["step"]
        ca = ch.category_axis
        ca.format.line.color.rgb = rgb(LINE_C)
        ca.tick_labels.font.size = Pt(18)
        ca.tick_labels.font.color.rgb = rgb(TEXT)
    return ch


LINE_C = "DDE2EC"


def L_chart(prs, s, n, where):
    """1–2 карточки с графиком (родные диаграммы PowerPoint — правятся в Excel) + плашка-вывод."""
    sl, top = light_slide(prs, s, where, n)
    charts = s["charts"]
    k = len(charts)
    g = 20
    note = s.get("note")
    nh = 150 if note else 0
    cw = (1800 - g * (k - 1)) / k
    ch_h = 1000 - top - (nh + g if note else 0)
    for i, c in enumerate(charts):
        x = 60 + i * (cw + g)
        rect(sl, x, top, cw, ch_h, CARD, radius=12)
        style = "grad" if i == 0 else "grey"
        card_head(sl, x, top, cw, 90, style, c.get("title", ""), c.get("icon"), size=28,
                  where=f"{where} график {i + 1}")
        chart(sl, x + 20, top + 110, cw - 40, ch_h - 130, c)
    if note:
        y = 1000 - nh
        rect(sl, 60, y, 1800, nh, CARD, radius=12)
        tx = 90
        if s.get("note_icon"):
            icon_pic(sl, s["note_icon"], 90, y + 24, 40)
            tx = 145
        head = s.get("note_title")
        if head:
            text(sl, tx, y + 20, 1700, 46, head, size=30, color=T["accent"], min_size=22)
        text(sl, 90, y + (72 if head else 20), 1740, nh - (90 if head else 40), note, size=22, color=TEXT,
             min_size=16, where=where + " вывод")
    if s.get("source"):
        text(sl, 60, 1010, 1500, 40, f"Источник: {s['source']}", size=16, color=GREY, min_size=16)
    return sl


def L_stats(prs, s, n, where):
    """Большие цифры (до 120 pt) с подписями в одну строку; note — вывод на плашке второго цвета снизу."""
    sl, top = light_slide(prs, s, where, n)
    items = s["items"]
    k = len(items)
    g = 20
    cw = (1800 - g * (k - 1)) / k
    ch = 382
    for i, it in enumerate(items):
        x = 60 + i * (cw + g)
        rect(sl, x, top, cw, ch, CARD, radius=12)
        unit = f"<g>{it['unit']}</g>" if it.get("unit") else ""
        text(sl, x + 40, top + 30, cw - 80, 150, f"{it['value']}{unit}", size=120, color=T["accent"], light=True,
             spacing=1.0, min_size=60, where=f"{where} цифра {i + 1}")
        text(sl, x + 40, top + 200, cw - 80, ch - 230, it.get("text", ""), size=28, color=TEXT, min_size=16,
             where=f"{where} цифра {i + 1}")
    y = top + ch + g
    if s.get("note"):
        rect(sl, 60, y, 1800, 1000 - y, CARD, radius=12)
        text(sl, 100, y + 40, 1720, 1000 - y - 80, s["note"], size=32, color=TEXT, anchor="m", min_size=20,
             where=where + " вывод")
    if s.get("source"):
        text(sl, 60, 1010, 1500, 40, f"Источник: {s['source']}", size=16, color=GREY, min_size=16)
    return sl


def L_timeline(prs, s, n, where):
    """Таймлайн: горизонтальная ось, точки по очереди сверху и снизу, год 48 pt и подпись в плашке."""
    sl, top = light_slide(prs, s, where, n)
    items = s["items"]
    k = len(items)
    ay = 610
    line(sl, 60, ay, 1860, ay, "DDE2EC", 2)
    bw_max = 330
    step = (1800 - bw_max - 60) / max(1, k - 1) if k > 1 else 0
    for i, it in enumerate(items):
        x = 120 + i * step
        up = i % 2 == 0
        y_end = ay - 150 if up else ay + 150
        line(sl, x, ay, x, y_end, MEDIUM, 2)
        dot = sl.shapes.add_shape(MSO_SHAPE.OVAL, E(x - 7), E(ay - 7), E(14), E(14))
        fill_solid(dot, MEDIUM)
        no_line(dot)
        bw = min(bw_max, 2 * step - 20) if k > 1 else 400
        by = y_end - 210 if up else y_end + 10
        text(sl, x, by, bw, 60, it.get("label", ""), size=48, color=T["accent"], spacing=1.0, min_size=28,
             where=f"{where} точка {i + 1}")
        rect(sl, x, by + 70, bw, 130, CARD, radius=8)
        text(sl, x + 16, by + 82, bw - 32, 106, it.get("text", ""), size=22, color=TEXT, min_size=16,
             where=f"{where} точка {i + 1}")
    return sl


def L_scheme(prs, s, n, where):
    """Структурная схема: 1–2 верхних блока → 2–4 элемента → итоговая плашка во всю ширину."""
    sl, top = light_slide(prs, s, where, n)
    tops, mids, bottom = s.get("top", []), s.get("items", []), s.get("bottom")
    g = 20
    y = top + 40
    th = 190
    if tops:
        k = len(tops)
        w = (1800 - g * (k - 1)) / k
        for i, b in enumerate(tops):
            x = 60 + i * (w + g)
            if i == 0:
                rect(sl, x, y, w, th, grad=(T["grad"][0], T["grad"][1], (T["grad"][0], 0.35)), radius=12)
                c1, c2 = WHITE, WHITE
            else:
                rect(sl, x, y, w, th, CARD, radius=12)
                c1, c2 = T["accent"], TEXT
            text(sl, x + 26, y + 18, w - 52, 44, b.get("title", ""), size=30, color=c1, min_size=20)
            text(sl, x + 26, y + 70, w - 52, th - 88, b.get("text", ""), size=22, color=c2, min_size=16,
                 where=f"{where} блок {i + 1}")
            line(sl, x + w / 2, y + th, x + w / 2, y + th + 40, "AEB8C8", 2)
        y += th + 40
    mh = 230
    if mids:
        k = len(mids)
        w = (1800 - g * (k - 1)) / k
        for i, b in enumerate(mids):
            x = 60 + i * (w + g)
            hl = b.get("highlight")
            if hl:
                rect(sl, x, y, w, mh, grad=(T["grad"][0], T["grad"][1], (T["grad"][0], 0.3)), radius=12)
            else:
                rect(sl, x, y, w, mh, CARD, radius=12)
            tx = x + 24
            if b.get("num"):
                text(sl, tx, y + 16, 90, 60, b["num"], size=48, color=WHITE if hl else T["accent"], light=True,
                     spacing=1.0, min_size=36)
                tx += 80
            elif b.get("icon"):
                icon_pic(sl, b["icon"], tx, y + 22, 40, white=hl)
                tx += 56
            text(sl, tx, y + 22, x + w - tx - 20, 50, b.get("title", ""), size=28, color=WHITE if hl else T["accent"],
                 min_size=18)
            text(sl, x + 24, y + 90, w - 48, mh - 106, b.get("text", ""), size=20, color=WHITE if hl else TEXT,
                 min_size=16, where=f"{where} элемент {i + 1}")
            line(sl, x + w / 2, y + mh, x + w / 2, y + mh + 40, "AEB8C8", 2)
        y += mh + 40
    if bottom:
        bh = max(110, 1000 - y)
        bh = min(bh, 140)
        rect(sl, 60, y, 1800, bh, grad=(T["grad"][0], T["grad"][1], (T["grad"][0], 0.3)), radius=12)
        text(sl, 100, y + 10, 1720, bh - 20, bottom, size=26, color=WHITE, align="c", anchor="m", min_size=16,
             where=where + " итог")
    return sl


def L_contacts(prs, s, n, where):
    """Контакты: карточки людей (ФИО 62 pt прописными, должность, телефон, почта) + QR справа."""
    s = {**s, "title": s.get("title", "Контакты")}
    sl, top = light_slide(prs, s, where, n)
    people = s["people"][:3]
    k = len(people)
    g = 20
    ch = (1000 - top - g * (k - 1)) / k
    for i, p in enumerate(people):
        y = top + i * (ch + g)
        first = i == 0
        has_qr = bool(p.get("qr"))
        cw = 1435 if has_qr else 1800
        if first:
            rect(sl, 60, y, cw, ch, grad=("30B6FD", "1B47CF", ("1E4BD2", 0.45)), radius=12)
        else:
            rect(sl, 60, y, cw, ch, CARD_HEAD, radius=12)
        c1 = WHITE if first else RETAIL
        c2 = WHITE if first else GREY
        text(sl, 100, y + 30, cw - 80, 80, p["name"], size=62 if k <= 2 else 48, color=c1, upper=True,
             spacing=1.0, min_size=36, where=f"{where} {i + 1}")
        text(sl, 100, y + 110, cw - 80, 56, p.get("role", ""), size=42 if k <= 2 else 32, color=c2, min_size=24)
        cy = y + ch - 70
        for j, (key, ic) in enumerate((("phone", "phone"), ("email", "mail"))):
            if p.get(key):
                xx = 100 + j * 520
                if (A / "icons" / f"{ic}.png").exists():
                    icon_pic(sl, ic, xx, cy, 40, white=first)
                text(sl, xx + 60, cy, 440, 44, p[key], size=24, color=c1, anchor="m", min_size=18)
        if has_qr:
            rect(sl, 1515, y, 345, ch, WHITE if first else CARD, radius=12, line=("DDE2EC", 1.5) if first else None)
            q = Path(p["qr"]) if Path(p["qr"]).is_absolute() else CTX["base"] / p["qr"]
            sz = min(255, ch - 90)
            picture(sl, q, 1515 + (345 - sz) / 2, y + 24, sz, sz)
            text(sl, 1535, y + ch - 60, 305, 40, p.get("qr_caption", ""), size=20, color=GREY, align="c",
                 min_size=16)
    return sl


def L_image(prs, s, n, where):
    """Фото во всю рабочую область с заголовком."""
    sl, top = light_slide(prs, s, where, n)
    picture(sl, asset("photo", s["photo"]), 60, top, 1800, 1000 - top, cover=True, radius=12)
    return sl


LAYOUTS = {k[2:]: v for k, v in globals().items() if k.startswith("L_")}


# --- Сборка ---------------------------------------------------------------
def build(deck, base, out):
    global T
    CTX["base"] = base
    CTX["direction"] = deck.get("direction", "vtb")
    if CTX["direction"] not in DIRECTIONS:
        sys.exit(f"Неизвестное направление «{CTX['direction']}». Доступны: {', '.join(DIRECTIONS)}")
    T = DIRECTIONS[CTX["direction"]]
    prs = Presentation()
    prs.slide_width, prs.slide_height = E(W), E(H)
    for i, s in enumerate(deck["slides"], 1):
        lay = s.get("layout", "bullets")
        if lay not in LAYOUTS:
            sys.exit(f"Слайд {i}: неизвестный макет «{lay}». Доступны: {', '.join(sorted(LAYOUTS))}")
        sl = LAYOUTS[lay](prs, s, i, f"Слайд {i} ({lay})")
        if s.get("notes"):
            sl.notes_slide.notes_text_frame.text = s["notes"]
    prs.core_properties.title = deck.get("title", "")
    prs.core_properties.author = "ВТБ Образование"
    p = out / "deck.pptx"
    prs.save(p)
    return p


FONTS_CONF = """<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd"><fontconfig>
<include ignore_missing="yes">/etc/fonts/fonts.conf</include><dir>{dir}</dir>
<alias binding="same"><family>{book}</family><prefer><family>Open Sans</family></prefer></alias>
<alias binding="same"><family>{light}</family><prefer><family>Open Sans Light</family></prefer></alias>
</fontconfig>"""


def preview(pptx, out):
    """PDF и PNG через LibreOffice. Если VTB Group не установлен — подставляется Open Sans."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice не найден — превью пропущено")
        return
    with tempfile.TemporaryDirectory() as td:
        conf = Path(td) / "fonts.conf"
        conf.write_text(FONTS_CONF.format(dir=A / "fonts", book=FONT, light=FONT_LIGHT))
        env = {**os.environ, "FONTCONFIG_FILE": str(conf), "HOME": td}
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out), str(pptx)],
                       env=env, check=True, capture_output=True, timeout=600)
    pdf = out / "deck.pdf"
    for f in out.glob("slide_*.png"):
        f.unlink()
    subprocess.run(["pdftoppm", "-png", "-scale-to-x", "1920", "-scale-to-y", "1080", str(pdf),
                    str(out / "slide")], check=True)
    for f in sorted(out.glob("slide-*.png")):
        f.rename(out / f"slide_{int(f.stem.split('-')[1]):02d}.png")


def main():
    global FONT, FONT_LIGHT
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("-o", "--out", default="out")
    ap.add_argument("--preview", action="store_true", help="сделать deck.pdf и slide_XX.png (нужен LibreOffice)")
    ap.add_argument("--font", default=FONT, help=f"шрифт основного текста (по умолчанию «{FONT}»)")
    ap.add_argument("--font-light", default=FONT_LIGHT, help=f"светлое начертание (по умолчанию «{FONT_LIGHT}»)")
    a = ap.parse_args()
    FONT, FONT_LIGHT = a.font, a.font_light
    src = Path(a.deck).resolve()
    deck = json.loads(src.read_text(encoding="utf-8"))
    out = Path(a.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    p = build(deck, src.parent, out)
    if a.preview:
        preview(p, out)
    (out / "report.json").write_text(json.dumps(REPORT, ensure_ascii=False, indent=1), encoding="utf-8")
    for r in REPORT:
        print("•", r)
    bad = [r for r in REPORT if "ПЕРЕПОЛНЕНИЕ" in r]
    print(f"\nГотово: {p}  (слайдов: {len(deck['slides'])}, замечаний: {len(REPORT)}, переполнений: {len(bad)})")
    if bad:
        sys.exit(2)


if __name__ == "__main__":
    main()
