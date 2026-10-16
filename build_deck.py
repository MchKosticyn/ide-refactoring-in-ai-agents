"""Сборка презентации доклада «Использование инструментов рефакторинга IDE в ИИ-агентах»
поверх фирменного шаблона Veai (берётся из вебинара 3).

Запуск:  python3 build_deck.py
Выход:   Доклад_ Рефакторинг в ИИ-агентах.pptx
"""
import copy
import os
import shutil

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

# Фирменный шаблон Veai (фон, макеты, шрифты) берётся из внутренней презентации вебинара.
# В репозиторий не входит. Путь можно переопределить переменной окружения VEAI_TEMPLATE.
TEMPLATE = os.environ.get(
    "VEAI_TEMPLATE",
    "/Users/michael/Documents/Work/webinars/Вебинар 3_ Работа с ИИ на уровне пользователя.pptx",
)
DECK = "Доклад_ Рефакторинг в ИИ-агентах.pptx"     # рабочий файл Михаила – скрипт его НЕ перезаписывает
OUT = os.path.join("build", DECK)                   # результат сборки; в рабочий файл копировать только после сравнения
os.makedirs("build", exist_ok=True)

TALK = "Использование инструментов рефакторинга IDE в ИИ-агентах"
TALK_SHORT = "Рефакторинг в ИИ-агентах"

# ---------- палитра шаблона ----------
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0xB8, 0xB8, 0xC8)
PURPLE = RGBColor(0x90, 0x66, 0xFF)
LILAC = RGBColor(0xC9, 0xB8, 0xFF)
RED = RGBColor(0xEA, 0x99, 0x99)
GREEN = RGBColor(0x93, 0xC4, 0x7D)
AMBER = RGBColor(0xF6, 0xC2, 0x6B)
FONT = "Golos Text"

# =====================================================================
# подготовка файла: копия шаблона, удаляем все слайды, оставляем макеты
# =====================================================================
shutil.copyfile(TEMPLATE, OUT)
prs = Presentation(OUT)
sldIdLst = prs.slides._sldIdLst
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
for sid in list(sldIdLst):
    prs.part.drop_rel(sid.get(RID))
    sldIdLst.remove(sid)
prs.save(OUT)
prs = Presentation(OUT)

LAYOUT = {l.name: l for l in prs.slide_masters[0].slide_layouts}
L_TEXT = LAYOUT["CUSTOM_19_1"]    # заголовок + большое текстовое поле
L_CARDS = LAYOUT["CUSTOM_20"]     # кикер + заголовок, дальше рисуем сами
L_TITLE = LAYOUT["BLANK_2"]       # титульный

slide_no = 0

# =====================================================================
# низкоуровневые помощники
# =====================================================================

def _clear_placeholders(slide, keep_idx=()):
    for shp in list(slide.placeholders):
        if shp.placeholder_format.idx not in keep_idx:
            shp._element.getparent().remove(shp._element)


def _style_runs(tf, size=None, color=WHITE, bold=None, italic=None, font=FONT):
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = font
            if size is not None:
                r.font.size = Pt(size)
            if color is not None:
                r.font.color.rgb = color
            if bold is not None:
                r.font.bold = bold
            if italic is not None:
                r.font.italic = italic


def textbox(slide, x, y, w, h, text, size=12, color=WHITE, bold=False,
            italic=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            line_spacing=None, space_after=None, font=FONT):
    """Текстовое поле. text — строка или список абзацев.
    Абзац может быть строкой или кортежем (текст, dict-стиль)."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    paras = text if isinstance(text, list) else [text]
    first = True
    for item in paras:
        st = {}
        if isinstance(item, tuple):
            item, st = item
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = st.get("align", align)
        if line_spacing or st.get("line_spacing"):
            p.line_spacing = st.get("line_spacing", line_spacing)
        if space_after is not None or "space_after" in st:
            p.space_after = Pt(st.get("space_after", space_after))
        r = p.add_run()
        r.text = item
        r.font.name = st.get("font", font)
        r.font.size = Pt(st.get("size", size))
        r.font.color.rgb = st.get("color", color)
        r.font.bold = st.get("bold", bold)
        r.font.italic = st.get("italic", italic)
    return tb


def bullets(slide, x, y, w, h, items, size=13, color=WHITE, gap=6):
    """Маркированный список. Элемент — строка, или (строка, уровень),
    или (строка, уровень, dict-стиль)."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    first = True
    for it in items:
        st = {}
        lvl = 0
        if isinstance(it, tuple):
            if len(it) == 3:
                it, lvl, st = it
            else:
                it, lvl = it
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(gap)
        _set_bullet(p)
        if lvl:
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(Emu(Inches(0.22 + 0.3 * lvl))))
        r = p.add_run()
        r.text = it
        r.font.name = FONT
        r.font.size = Pt(st.get("size", size - 1.5 * lvl))
        r.font.color.rgb = st.get("color", color if lvl == 0 else GREY)
        r.font.bold = st.get("bold", False)
        r.font.italic = st.get("italic", False)
    return tb


def _fill_tf(tf, paras, size, color, bold, italic, align, font, gap):
    """Заполняет text_frame абзацами. Абзац: строка | (строка, dict) | ("•", строка) для буллета."""
    first = True
    for item in paras:
        st = {}
        bullet = False
        if isinstance(item, tuple):
            if len(item) == 2 and item[0] == "•":
                bullet, item = True, item[1]
            else:
                item, st = item
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = st.get("align", align)
        if gap is not None or "space_after" in st:
            p.space_after = Pt(st.get("space_after", gap))
        if bullet:
            _set_bullet(p)
        r = p.add_run()
        r.text = item
        r.font.name = st.get("font", font)
        r.font.size = Pt(st.get("size", size))
        r.font.color.rgb = st.get("color", color)
        r.font.bold = st.get("bold", bold)
        r.font.italic = st.get("italic", italic)


def _set_bullet(p):
    """Нативный маркер списка (а не символ в тексте)."""
    from pptx.oxml.ns import qn
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Emu(Inches(0.22))))
    pPr.set("indent", str(-Emu(Inches(0.22))))
    for tag in ("a:buNone", "a:buFont", "a:buChar", "a:buAutoNum", "a:buClr"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    # порядок дочерних элементов важен: buClr, buFont, buChar
    clr = pPr.makeelement(qn("a:buClr"), {})
    srgb = clr.makeelement(qn("a:srgbClr"), {"val": "FFFFFF"})
    clr.append(srgb)
    pPr.append(clr)
    pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
    pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))


CORNER = 0.14   # радиус скругления, дюймы – один для всех карточек и картинок


def corner_adj(w, h):
    """adj для roundRect: радиус = adj × min(w, h). Фиксированный радиус независимо от размера фигуры."""
    return min(0.5, CORNER / min(w, h))


def card(slide, x, y, w, h, line=WHITE, width=2.25, fill=None, text=None, size=12,
         color=WHITE, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, gap=None,
         margin=0.2, font=FONT):
    """Скруглённая плашка в стиле шаблона. Текст – внутри самой фигуры.
    text – строка или список абзацев (см. _fill_tf)."""
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                               Inches(w), Inches(h))
    s.adjustments[0] = corner_adj(w, h)
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    s.line.color.rgb = line
    s.line.width = Pt(width)
    s.shadow.inherit = False
    tf = s.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = Inches(0.12)
    if text is None:
        tf.text = ""
    else:
        paras = text if isinstance(text, list) else [text]
        _fill_tf(tf, paras, size, color, bold, False, align, font, gap)
    return s


def arrow(slide, x, y, w=0.32, h=0.14, color=PURPLE):
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y),
                               Inches(w), Inches(h))
    a.fill.solid()
    a.fill.fore_color.rgb = color
    a.line.color.rgb = color
    a.line.width = Pt(0.75)
    a.shadow.inherit = False
    return a


def axis(slide, x, y, w, left="синтаксис", right="семантика", color=GREY, size=10, width=1.25):
    """Ось со стрелкой вправо и подписями концов: синтаксис → семантика."""
    from pptx.oxml.ns import qn as _q
    ln = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    ln.line._get_or_add_ln().append(ln.line._get_or_add_ln().makeelement(_q("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    textbox(slide, x, y + 0.05, w / 2, 0.25, left, size=size, color=color)
    textbox(slide, x + w / 2, y + 0.05, w / 2, 0.25, right, size=size, color=color, align=PP_ALIGN.RIGHT)
    return ln


def hline(slide, x, y, w, color=PURPLE, width=1.0):
    ln = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    return ln


def vline(slide, x, y, h, color=PURPLE, width=1.0):
    ln = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x), Inches(y + h))
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    return ln


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# =====================================================================
# типы слайдов
# =====================================================================

BG_GLOW = os.environ.get("BG_GLOW", "assets/bg-glow.gif")   # анимированный фон из вебинаров Veai; вне репозитория


def bg_glow(slide):
    """Анимированный фон на весь слайд, под всеми фигурами."""
    if not os.path.exists(BG_GLOW):
        return None
    pic = slide.shapes.add_picture(BG_GLOW, 0, 0, width=prs.slide_width, height=prs.slide_height)
    tree = slide.shapes._spTree
    tree.remove(pic._element)
    tree.insert(2, pic._element)
    return pic


def _chrome(slide, kicker):
    """Кикер слева сверху и номер слайда справа — как в шаблоне."""
    global slide_no
    slide_no += 1
    textbox(slide, 0.4, 0.24, 4.74, 0.26, kicker, size=12, color=GREY)
    textbox(slide, 9.0, 0.24, 0.6, 0.26, str(slide_no), size=10, color=GREY,
            align=PP_ALIGN.RIGHT)


def slide_title(title_lines, speaker, subtitle=None):
    global slide_no
    slide_no += 1
    s = prs.slides.add_slide(L_TITLE)
    _clear_placeholders(s)
    bg_glow(s)
    s.shapes.add_picture("veai_logo.png", Inches(0.4), Inches(0.4), width=Inches(1.01))
    textbox(s, 0.4, 2.78, 9.0, 0.53, speaker, size=14, color=GREY)
    textbox(s, 0.4, 3.32, 9.2, 2.0, title_lines, size=36, color=WHITE, bold=False,
            anchor=MSO_ANCHOR.TOP, line_spacing=1.0)
    if subtitle:
        textbox(s, 0.4, 2.55, 9.0, 0.4, subtitle, size=14, color=LILAC)
    textbox(s, 9.0, 0.24, 0.6, 0.26, str(slide_no), size=10, color=GREY,
            align=PP_ALIGN.RIGHT)
    return s


def slide_section(number, title, sub=None):
    """Разделитель: большой номер справа, заголовок слева (как слайд 5 шаблона)."""
    s = prs.slides.add_slide(L_TEXT)
    _clear_placeholders(s)
    _chrome(s, TALK_SHORT)
    textbox(s, 0.4, 1.5, 6.6, 2.4, title, size=38, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if sub:
        textbox(s, 0.4, 4.05, 6.6, 0.8, sub, size=15, color=GREY)
    textbox(s, 7.3, 0.9, 2.3, 3.6, str(number), size=110, color=PURPLE,
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    return s


def slide_text(kicker, title, items, size=14, title_size=26, note=None):
    """Заголовок + маркированный список (как слайд 4 шаблона)."""
    s = prs.slides.add_slide(L_TEXT)
    _clear_placeholders(s)
    _chrome(s, kicker)
    textbox(s, 0.4, 0.55, 9.2, 0.7, title, size=title_size, color=WHITE)
    bullets(s, 0.45, 1.4, 9.1, 3.9, items, size=size, gap=8)
    if note:
        notes(s, note)
    return s


def slide_blank(kicker, title, title_size=26):
    """Кикер + заголовок, остальное рисуем вручную."""
    s = prs.slides.add_slide(L_TEXT)
    _clear_placeholders(s)
    _chrome(s, kicker)
    textbox(s, 0.4, 0.55, 9.2, 0.7, title, size=title_size, color=WHITE)
    return s


def slide_statement(kicker, big, small=None, note=None):
    """Один крупный тезис по центру."""
    s = prs.slides.add_slide(L_TEXT)
    _clear_placeholders(s)
    _chrome(s, kicker)
    card(s, 0.7, 1.6, 8.6, 2.2, line=PURPLE)
    textbox(s, 0.9, 1.7, 8.2, 2.0, big, size=26, color=WHITE, bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if small:
        textbox(s, 0.9, 4.0, 8.2, 1.0, small, size=14, color=GREY,
                align=PP_ALIGN.CENTER)
    if note:
        notes(s, note)
    return s


def slide_todo(kicker, title, what, items):
    """Заглушка под внутренние детали / демо. Жёлтая рамка, чтобы не потерять."""
    s = slide_blank(kicker, title)
    card(s, 0.5, 1.35, 9.0, 3.85, line=AMBER, size=13.5, gap=8,
         text=[("TODO: " + what, {"size": 13, "color": AMBER, "bold": True, "space_after": 10})]
              + [("•", it if isinstance(it, str) else it[0]) for it in items])
    return s


def slide_demo(kicker, title, run, steps):
    """Демо: видео 16:9 слева, справа карточка «Агент / Модель»; шаги – в заметки."""
    s = slide_blank(kicker, title)
    textbox(s, 0.49, 1.17, 9.0, 0.3, f"{run['agent']} · {run['model']}", size=11.5, color=GREY)
    w = 7.2; h = w * 9 / 16
    s.shapes.add_movie(run["video"], Inches((10 - w) / 2), Inches(1.5), Inches(w), Inches(h),
                       poster_frame_image=run["poster"], mime_type="video/mp4")
    notes(s, run["meta"] + ". Что происходит в ролике: " + " → ".join(steps) + ".")
    return s


RESULT_ASPECTS = ["Полнота", "Корректность"]


def result_table(s, statuses, metrics=None, banner=None):
    """Разбор прохода: полнота и корректность + полоса метрик.
    statuses – список (ok: True/False/None, 2–4 слова); None – жёлтый «по факту». Пояснения – словами."""
    y = 1.4
    for name, (ok, text) in zip(RESULT_ASPECTS, statuses):
        col = AMBER if ok is None else (GREEN if ok else RED)
        sym = "◐" if ok is None else ("✓" if ok else "✗")
        textbox(s, 0.49, y, 2.4, 1.0, name, size=16, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        if not text:
            card(s, 2.95, y, 6.55, 1.0, line=col, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER,
                 text=[(sym, {"size": 26, "bold": True, "color": col})])
            y += 1.2
            continue
        card(s, 2.95, y, 6.55, 1.0, line=col, anchor=MSO_ANCHOR.MIDDLE, margin=0.25,
             text=[(f"{sym}  {text}", {"size": 15})], color=WHITE)
        p = s.shapes[-1].text_frame.paragraphs[0]
        r0 = p.runs[0]
        full = r0.text
        r0.text = sym
        r0.font.size = Pt(22); r0.font.bold = True; r0.font.color.rgb = col
        r1 = p.add_run(); r1.text = full[len(sym):]
        r1.font.name = FONT; r1.font.size = Pt(15); r1.font.color.rgb = todo_color(text, WHITE)
        y += 1.2
    if metrics:
        n = len(metrics)
        w = 9.01 / n if n > 1 else 3.0
        x0 = 0.49 if n > 1 else 3.5          # одна метрика – узкая карточка по центру
        for i, (k, v) in enumerate(metrics):
            card(s, x0 + i * w, 4.0, w - 0.1, 1.0, line=PURPLE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
                 text=[(k, {"size": 10, "color": LILAC}), (v, {"size": 16, "bold": True, "color": todo_color(v)})])
    if banner:
        bottom_banner(s, banner, y=4.75, size=13)


def slide_pros_cons(kicker, pros, cons, note=None, y=1.4, h=3.7):
    s = slide_blank(kicker, "Плюсы и минусы")
    card(s, 0.49, y, 4.4, h, line=GREEN, size=12, gap=7,
         text=[("Плюсы", {"size": 17, "bold": True, "space_after": 12})] + [("•", p) for p in pros])
    card(s, 5.1, y, 4.4, h, line=RED, size=12, gap=7,
         text=[("Минусы", {"size": 17, "bold": True, "space_after": 12})] + [("•", c) for c in cons])
    anim(s, [0], [1])
    if note:
        notes(s, note)
    return s


PLAN = ["Введение", "Подходы к рефакторингу в агентах", "Итоги"]
APPROACHES = ["Текстовый подход", "LSP-серверы", "Инструменты IDE"]


def slide_plan(done=(), current=None):
    """План доклада перед каждой главой: главы в ряд с номерами, на анимированном фоне.
    done – пройденные (зелёным), current – текущая (белым, жирным), остальные – серым."""
    s = prs.slides.add_slide(L_TEXT)
    _clear_placeholders(s)
    bg_glow(s)
    _chrome(s, TALK_SHORT)
    textbox(s, 0.4, 0.55, 9.2, 0.7, "План доклада", size=26, color=WHITE)
    n = len(PLAN)
    x0, w = 0.6, 8.8 / n
    for i, title in enumerate(PLAN):
        x = x0 + i * w
        col = GREEN if title in done else (WHITE if title == current else GREY)
        textbox(s, x, 1.9, w - 0.3, 0.7, f"{i + 1:02d}", size=30, color=PURPLE if title == current else col, bold=title == current)
        textbox(s, x, 2.7, w - 0.3, 1.4, title, size=15, color=col, bold=title == current)
        if i:
            vline(s, x - 0.15, 1.95, 2.0, color=GREY, width=0.75)
    return s


def slide_subplan(current=None, done=()):
    """План главы «Подходы»: показываются только пройденные (зелёные) и текущий (фиолетовый); будущие – нет."""
    s = prs.slides.add_slide(L_TEXT)
    _clear_placeholders(s)
    bg_glow(s)
    _chrome(s, TALK_SHORT)
    textbox(s, 0.4, 0.55, 9.2, 0.7, "Подходы к рефакторингу в агентах", size=26, color=WHITE)
    geo = [(0.67, 2.49), (3.79, 2.41), (6.92, 2.31)]   # x, w – раскладка Михаила
    for i, title in enumerate(APPROACHES):
        if title != current and title not in done:
            break
        x, w = geo[i]
        col = GREEN if title in done else PURPLE
        card(s, x, 1.9, w, 1.5, line=col, width=2.25 if title == current else 1.5, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER,
             text=[(str(i + 1), {"size": 26, "bold": True, "color": col}), (title, {"size": 15, "bold": title == current, "color": WHITE if title == current else col})])
        if i < 2 and (APPROACHES[i + 1] == current or APPROACHES[i + 1] in done):
            gap_c = (x + w + geo[i + 1][0]) / 2
            arrow(s, gap_c - 0.2, 2.5, w=0.4, h=0.3, color=GREY)
    axis(s, geo[0][0], 3.75, geo[2][0] + geo[2][1] - geo[0][0], size=13, color=WHITE, width=2.0)
    return s


X3 = [0.49, 3.57, 6.65]


def three_cards(s, titles, bodies, y=1.45, h=2.7, numbered=True, line=WHITE):
    xs = X3
    for i, (t, b) in enumerate(zip(titles, bodies)):
        paras = []
        if numbered:
            paras.append((str(i + 1), {"size": 34, "color": PURPLE, "bold": True, "space_after": 2}))
        paras.append((t, {"size": 15, "bold": True, "space_after": 8}))
        paras.append((b, {"size": 11.5}))
        card(s, xs[i], y, 2.85, h, line=line if not isinstance(line, list) else line[i],
             text=paras, margin=0.25)


def todo_color(text, default=WHITE):
    return AMBER if "TODO" in text else default


def bottom_banner(s, text, y=4.4, size=15, color=None):
    card(s, 1.2, y, 7.6, 0.68, line=PURPLE, text=text, size=size, bold=True,
         color=color or todo_color(text), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)


X4 = [0.49, 2.78, 5.07, 7.36]       # четыре карточки в ряд
X4N = [0.72, 3.01, 5.33, 7.61]      # четыре узкие карточки со стрелками


COMPARE_COLS = ["Человек с IDE", "Агент + текстовый подход", "Агент + LSP", "Агент + инструменты IDE"]
COMPARE_ROWS = [
    ("Полнота", ["✓", "✓", "✓", "✓"]),
    ("Корректность", ["✓", "✗", "✓", "✓"]),
    ("Время", ["≈ 40 мин", "6:20", "4:38", "4:54"]),
    ("Ошибок / Исправлено", ["–", "6 / 3", "3 / 3", "2 / 2"]),
    ("Токены", ["–", "1,8M", "2,4M", "3,3M"]),
    ("Стартовый контекст", ["–", "49k", "59k", "≈ 20k"]),
]


def compare_table(s, upto=None):
    """Таблица «Результат и цена». upto – последняя показанная колонка (таблица растёт по ходу доклада)."""
    n = COMPARE_COLS.index(upto) + 1 if upto else len(COMPARE_COLS)
    cols = [""] + COMPARE_COLS[:n]
    x0, y0 = 0.49, 1.35
    cw = [2.3] + [6.7 / n] * n
    rh = 0.43
    x = x0
    for i, c in enumerate(cols):
        if c:
            textbox(s, x, y0, cw[i], 0.4, c, size=11 if n > 2 else 12, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
        x += cw[i]
    hline(s, x0, y0 + 0.42, sum(cw), color=PURPLE, width=1.2)
    colmap = {"✓": GREEN, "✗": RED}
    for r, (name, vals) in enumerate(COMPARE_ROWS):
        y = y0 + 0.5 + r * rh
        textbox(s, x0, y, cw[0], rh, name, size=11.5, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        x = x0 + cw[0]
        for i, v in enumerate(vals[:n]):
            col = colmap.get(v[0], todo_color(v))
            textbox(s, x, y, cw[i + 1], rh, v, size=11.5, color=col, bold=v[0] in colmap, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            x += cw[i + 1]
        if r == 1:
            hline(s, x0, y + rh, sum(cw), color=GREY, width=0.5)


def two_columns(s, left_title, left_items, right_title, right_items,
                left_color=RED, right_color=GREEN, y=1.35, h=3.75,
                left_kicker=None, right_kicker=None):
    def col(x, color, kicker, title, items):
        paras = []
        if kicker:
            paras.append((kicker, {"size": 10.5, "color": color, "bold": True, "space_after": 4}))
        paras.append((title, {"size": 16, "bold": True, "space_after": 10}))
        paras += [("•", it) for it in items]
        card(s, x, y, 4.4, h, line=color, size=12, gap=5, text=paras)
    col(0.49, left_color, left_kicker, left_title, left_items)
    col(5.1, right_color, right_kicker, right_title, right_items)




def picture(s, path, x, y, w=None, h=None):
    """Картинка по ширине или высоте с сохранением пропорций, со скруглением как у карточек."""
    from pptx.oxml.ns import qn as _q
    if w is not None:
        pic = s.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w))
    else:
        pic = s.shapes.add_picture(path, Inches(x), Inches(y), height=Inches(h))
    geom = pic._element.spPr.find(_q("a:prstGeom"))
    geom.set("prst", "roundRect")
    av = geom.find(_q("a:avLst"))
    for gd in list(av):
        av.remove(gd)
    adj = corner_adj(Emu(pic.width).inches, Emu(pic.height).inches)
    av.append(av.makeelement(_q("a:gd"), {"name": "adj", "fmla": f"val {int(adj * 100000)}"}))
    return pic


def source_line(s, x, y, w, text):
    """Подпись источника над картинкой – одна строка серым."""
    textbox(s, x, y, w, 0.3, text, size=11, color=GREY)


# ---------- анимация: появление по клику ----------
from lxml import etree as _et
_NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
       "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

def animate(slide, groups):
    """Появление по клику. groups – список групп; группа – список фигур, появляющихся одним кликом."""
    groups = [g for g in groups if g]
    if not groups:
        return
    tid = [1]
    def nid():
        tid[0] += 1
        return str(tid[0])
    def effect(spid, node_type):
        return (f'<p:par><p:cTn id="{nid()}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="{node_type}">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                f'<p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>')
    seqs = []
    for g in groups:
        inner = "".join(effect(sh.shape_id, "clickEffect" if i == 0 else "withEffect") for i, sh in enumerate(g))
        seqs.append(f'<p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                    f'<p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{inner}'
                    f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    xml = f'''<p:timing xmlns:p="{_NS["p"]}" xmlns:a="{_NS["a"]}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
<p:seq concurrent="1" nextAc="seek"><p:cTn id="{nid()}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{"".join(seqs)}</p:childTnLst></p:cTn>
<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>
</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'''
    el = _et.fromstring(xml)
    old = slide._element.find(_qn2("p:timing"))
    if old is not None:
        slide._element.remove(old)
    slide._element.append(el)


def elbow_arrow(slide, points, r=0.18, color=GREEN, width=1.5, dashed=True, head=True):
    """Ортогональная ломаная через points (дюймы) со скруглёнными углами радиуса r; стрелка на конце."""
    from pptx.oxml.ns import qn as _q
    import math
    xs = [p[0] for p in points]; ys = [p[1] for p in points]
    x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
    pad = 0.05
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x0 - pad), Inches(y0 - pad), Inches(x1 - x0 + 2 * pad), Inches(y1 - y0 + 2 * pad))
    E = lambda v: int(round(v * 914400))
    W, H = E(x1 - x0 + 2 * pad), E(y1 - y0 + 2 * pad)
    P = lambda x, y: (E(x - x0 + pad), E(y - y0 + pad))
    cmds = []
    px, py = P(*points[0]); cmds.append(f'<a:moveTo><a:pt x="{px}" y="{py}"/></a:moveTo>')
    for i in range(1, len(points) - 1):
        (ax, ay), (bx, by), (cx, cy) = points[i - 1], points[i], points[i + 1]
        d1 = (math.copysign(1, bx - ax) if bx != ax else 0, math.copysign(1, by - ay) if by != ay else 0)
        d2 = (math.copysign(1, cx - bx) if cx != bx else 0, math.copysign(1, cy - by) if cy != by else 0)
        A = (bx - d1[0] * r, by - d1[1] * r)
        lx, ly = P(*A); cmds.append(f'<a:lnTo><a:pt x="{lx}" y="{ly}"/></a:lnTo>')
        st = (math.degrees(math.atan2(-d2[1], -d2[0])) % 360) * 60000
        cross = d1[0] * d2[1] - d1[1] * d2[0]
        sw = 5400000 if cross > 0 else -5400000
        cmds.append(f'<a:arcTo wR="{E(r)}" hR="{E(r)}" stAng="{int(st)}" swAng="{sw}"/>')
    lx, ly = P(*points[-1]); cmds.append(f'<a:lnTo><a:pt x="{lx}" y="{ly}"/></a:lnTo>')
    geom = _et.fromstring(f'<a:custGeom xmlns:a="{_NS["a"]}"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
                          f'<a:pathLst><a:path w="{W}" h="{H}" fill="none">{"".join(cmds)}</a:path></a:pathLst></a:custGeom>')
    spPr = sh._element.spPr
    spPr.replace(spPr.find(_q("a:prstGeom")), geom)
    sh.fill.background()
    sh.line.color.rgb = color
    sh.line.width = Pt(width)
    sh.shadow.inherit = False
    ln = sh.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(_q("a:prstDash"), {"val": "dash"}))
    ln.append(ln.makeelement(_q("a:round"), {}))
    if head:
        ln.append(ln.makeelement(_q("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return sh


ANIM = {}


def anim(slide, *groups):
    """Задать анимацию слайда: группы индексов содержательных фигур (0 – первая после кикера, номера и заголовка).
    Слайды без anim() появляются целиком: демо, скриншоты, таблицы, схемы."""
    ANIM[slide.slide_id] = [list(g) for g in groups]


def _qn2(tag):
    pfx, name = tag.split(":")
    return "{%s}%s" % (_NS[pfx], name)


# =====================================================================
# СЛАЙДЫ  (по «План подробный.md»)
# =====================================================================

CH_INTRO = "Введение"
CH_APPR = "Подходы к рефакторингу в агентах"
MONO = "Menlo"
DEMO = {   # финальные прогоны Максима (demo-handoff, 2026-10-15); «Попытки» – ошибок по ходу / исправил сам
    "text":   dict(video="demo-handoff/runs/1-claude-code/session-edited.mp4", poster="img/demo-poster-1.png",
                   agent="Claude Code", model="Opus 5.5", meta="Прогон 6:20, ролик 2:38",
                   metrics=[("Время", "6:20"), ("Ошибок / Исправлено", "6 / 3"), ("Токены", "1,8M"), ("Стартовый контекст", "49k")]),
    "lsp":    dict(video="demo-handoff/runs/2-claude-code-serena/session-edited.mp4", poster="img/demo-poster-2.png",
                   agent="Claude Code + Serena (LSP)", model="Opus 5.5", meta="Прогон 4:38, ролик 2:42",
                   metrics=[("Время", "4:38"), ("Ошибок / Исправлено", "3 / 3"), ("Токены", "2,4M"), ("Стартовый контекст", "59k")]),
    "engine": dict(video="demo-handoff/runs/3-veai/session-edited.mp4", poster="img/demo-poster-3.png",
                   agent="Veai", model="Opus 5.5", meta="Прогон 4:54, ролик 2:20",
                   metrics=[("Время", "4:54"), ("Ошибок / Исправлено", "2 / 2"), ("Токены", "3,3M"), ("Стартовый контекст", "≈ 20k")]),
}

# ---- Титул ----
s = slide_title(["Использование инструментов", "рефакторинга IDE в ИИ-агентах"],
                ["Спикер: Михаил Костицын", "Эксперт: Максим Паршин"])

# ---- О нас ----
s = slide_blank(TALK_SHORT, "О нас")
people = [("Михаил Костицын", "Спикер", "img/photo-michael.jpg",
           ["Ведущий разработчик Veai",
            "Начал путь в разработке с анализа программ и формальной верификации: JetBrains Research, Huawei",
            "С 2023 года занимаюсь разработкой систем с искусственным интеллектом"]),
          ("Максим Паршин", "Эксперт", "img/photo-maxim.jpg",
           ["Разработчик Veai",
            "Ранее разрабатывал движок символьного исполнения и генерацию тестов в Huawei",
            "Сейчас – формальные методы анализа программ вместе с ИИ-моделями"])]
for i, (name, role, photo, lines) in enumerate(people):
    x = 0.49 + i * 4.6
    card(s, x, 1.45, 4.4, 3.7, line=PURPLE if i == 0 else WHITE)
    picture(s, photo, x + 0.25, 1.7, w=1.5)
    textbox(s, x + 1.95, 1.75, 2.3, 0.53, [(name, {"size": 16, "bold": True}), (role, {"size": 12, "color": LILAC})], size=16)
    textbox(s, x + 0.25, 3.35, 3.9, 1.44, [(ln, {"color": todo_color(ln, WHITE)}) for ln in lines], size=12, color=WHITE, space_after=6)

# ---- план ----
slide_plan(current=CH_INTRO)

# ---- Рамка ----
s = slide_blank(CH_INTRO, "Определение")
card(s, 0.72, 1.77, 8.56, 0.81, line=PURPLE, text="Рефакторинг – это изменение структуры кода без изменения поведения",
     size=15, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
card(s, 0.72, 2.88, 4.1, 1.39, line=WHITE, size=12.5, gap=8, text=[
    ("Задача 1: найти все места", {"size": 15, "bold": True, "space_after": 10}),
    "Нужна полнота: 97 % не годится",
    ("78 вызовов × 97 % = 2 пропуска", {"color": LILAC, "bold": True}),
])
card(s, 5.18, 2.88, 4.1, 1.39, line=WHITE, size=12.5, gap=8, text=[
    ("Задача 2: применить", {"size": 15, "bold": True, "space_after": 10}),
    "Нужна корректность: правка не меняет поведение программы",
])
anim(s, [0], [1], [2])
notes(s, "Каждый проход дальше разбираем одинаково: найдены ли все места, только ли нужные правки, сохранено ли поведение – и цена: время, стоимость, ходы, токены, стартовый контекст.")

# ---- Зачем это QA ----
s = slide_blank(CH_INTRO, "Зачем это QA")
card(s, 0.49, 1.6, 9.0, 2.25, line=WHITE, size=16, gap=12, margin=0.35, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("•", "Миграция фреймворка: Selenium → Playwright, JUnit 4 → 5"),
    ("•", "Рефакторинг тестов: переход на Page Object, унификация, фикстуры"),
    ("•", "Смена тестового API: переименование, смена базового класса тестов"),
    ("•", "…"),
])
card(s, 0.49, 4.1, 9.0, 0.9, line=RED, size=15, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, text=
     "Зелёные тесты не всегда доказывают, что поведение не изменилось")
anim(s, [0], [1])

# ---- Кейс ----
s = slide_blank(CH_INTRO, "Кейс: detekt")
textbox(s, 0.49, 1.25, 9.0, 0.48, [
    "detekt – статический анализатор Kotlin: набор правил, каждое ищет в коде свой вид проблем",
    "Правила тестируют на сниппетах – короткие фрагменты Kotlin-кода прямо в тесте"], size=12.5, color=GREY)
card(s, 0.49, 1.97, 9.0, 1.15, line=WHITE, size=11.5, font=MONO, gap=1, margin=0.15, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("fun Rule.compileAndLint(code)", {"color": LILAC, "bold": True}),
    ("    сначала компилирует сниппет компилятором Kotlin, потом прогоняет правило", {"color": GREY, "font": FONT}),
    ("fun Rule.lint(code)", {"color": LILAC, "bold": True}),
    ("    сразу прогоняет правило: сниппет может быть неполным и не компилироваться", {"color": GREY, "font": FONT}),
])
card(s, 0.49, 3.39, 9.0, 1.77, line=PURPLE, size=12, gap=3, text=[
    ("ЗАДАЧА: PR #7873", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 6}),
    ("Оставить одну функцию: lint(code, compile = true)", {"bold": True}),
    "compileAndLint(code) → lint(code) – 1428 вызовов",
    "lint(code) → lint(code, compile = false) – 78 вызовов",
    ("Суммарно нужно изменить: 1506 вызовов", {"bold": True}),
    ("Не менять: перегрузку lint(ktFile) – 36 вызовов, FormattingRule.lint – 66 вызовов", {"bold": True}),
])
anim(s, [0, 1], [2])
notes(s, "«Только в CI»: compileAndLint компилирует сниппет, только если сборка запущена с -Pcompile-test-snippets=true; это делает отдельный job в GitHub Actions, "
         "локально флаг выключен и обе функции ведут себя одинаково. В терминах detekt автор PR упрощает API тестовой библиотеки detekt-test: "
         "вместо пары lint / compileAndLint – одна lint с параметром compile. Ловушки не называем – всплывут в разборах.")

# ---- Человек в IDE: PR ----
s = slide_blank(CH_INTRO, "Кейс: detekt")
picture(s, "img/pr-7873-header.png", 0.58, 1.39, w=8.85)              # 1560×650 → 8.85×3.69
notes(s, "Два коммита за 20 минут, PR открыт через 8 минут после второго – около 30 минут работы. На GitHub у коммитов стоит 19 февраля – это дата rebase перед слиянием; авторские даты – 3 января 2025, 11:12 и 11:32, PR открыт в 11:40. "
         "Что делал автор (по диффу PR): Change Signature на старом lint – параметр compile, в вызовах false; Rename compileAndLint → lint; тела слил вручную. "
         "134 файла, +1693 −1726.")

# ---- Инструменты рефакторинга IDE: gif ----
ide_tools = [("Find Usages", "img/ide-find-usages.gif", "⌥F7 – все использования compileAndLint: 1428 вызовов и 130 импортов"),
             ("Rename", "img/ide-rename.gif", "⇧F6 – compileAndLint → lint: конфликт «функция lint уже объявлена» – до применения"),
             ("Change Signature", "img/ide-change-signature.gif", "⌘F6 – параметр compile: Boolean = false в старом lint: 76 вызовов обновлены"),
             ("Safe Delete", "img/ide-safe-delete.gif", "⌘⌫ – удалить compileAndLint: 1428 использований, удаление заблокировано")]
for name, gif, sub in ide_tools:
    s = slide_blank(CH_INTRO, f"Инструменты IDE: {name}")
    from PIL import Image as _Img
    _w, _h = _Img.open(gif).size
    gw = min(7.2, 4.11 * _w / _h)                                      # раскладка Михаила: 7.2×4.11, y 1.3
    picture(s, gif, (10 - gw) / 2, 1.3, w=gw)
    notes(s, sub + ". Запись IntelliJ IDEA на detekt (2026-10-14), gif зациклен. Говорим коротко: что нажали и что IDE сделала сама.")

# ---- Человек в IDE: разбор ----
s = slide_blank(CH_INTRO, "Человек с IDE: результаты")
result_table(s, [
    (True, "все 1506 вызовов"),
    (True, "поведение не изменилось"),
], metrics=[("Время", "≈ 40 мин")])
textbox(s, 0.49, 5.1, 9.0, 0.35, "Можно ли быстрее?", size=14, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
anim(s, [0, 1], [2, 3], [4], [5])
notes(s, "Rename в занятое имя – IDE останавливает и показывает конфликт. Это то, что будем искать у агентов. "
         "Время – по ускоренной записи с таймером. Порядок величины подтверждает история PR: оба коммита автора сделаны за 20 минут 3 января 2025 (11:12 и 11:32). "
         "Стоимость не сравниваем: у агентов она в логах не записана, тарифы и подписки разные.")

# ---- Термины ----
s = slide_blank(CH_INTRO, "ИИ-агент")
card(s, 0.49, 1.89, 2.9, 2.25, line=WHITE, size=11.5, gap=4, text=[
    ("КОНТЕКСТ", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 2}),
    ("что модель видит", {"size": 12, "color": GREY, "space_after": 12}),
    ("•", "системный промпт"), ("•", "история чата"), ("•", "прочитанные файлы"), ("•", "результаты вызовов инструментов"),
])
card(s, 3.75, 2.03, 2.5, 1.2, line=PURPLE, fill=RGBColor(0x1A, 0x12, 0x33), align=PP_ALIGN.CENTER,
     anchor=MSO_ANCHOR.MIDDLE, text=[("МОДЕЛЬ", {"size": 14, "bold": True}),
                                     ("генерирует сообщение\nили вызов инструмента", {"size": 10.5, "color": GREY})])
arrow(s, 3.42, 2.56, w=0.3, h=0.14)
arrow(s, 6.28, 2.56, w=0.3, h=0.14)
card(s, 6.61, 1.89, 2.9, 2.25, line=WHITE, size=11.5, gap=4, text=[
    ("ИНСТРУМЕНТЫ", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 2}),
    ("что модель может сделать", {"size": 12, "color": GREY, "space_after": 12}),
    ("•", "чтение файла"), ("•", "поиск по тексту"), ("•", "правка файла"), ("•", "запуск тестов"),
])
ret = s.shapes.add_shape(MSO_SHAPE.LEFT_ARROW, Inches(3.42), Inches(3.74), Inches(3.16), Inches(0.16))
ret.fill.solid(); ret.fill.fore_color.rgb = LILAC; ret.line.color.rgb = LILAC; ret.shadow.inherit = False
textbox(s, 3.42, 3.92, 3.16, 0.3, "результат вызова", size=10.5, color=GREY, align=PP_ALIGN.CENTER)
textbox(s, 2.55, 4.84, 4.88, 0.25,
        "ИИ-агент = модель + контекст + инструменты",
        size=11.5, color=GREY, align=PP_ALIGN.CENTER)

# ---- план ----
slide_plan(done={CH_INTRO}, current=CH_APPR)
s = slide_subplan(current="Текстовый подход")
notes(s, "Ось: синтаксис – как код записан: буквы, строки, имена. Семантика – что он значит: какая функция вызывается на самом деле. "
         "subject.lint(code) в двух файлах пишется одинаково, но может вызывать разные функции. grep видит запись, IDE видит смысл. "
         "Три подхода – три точки на этой оси, слева направо.")

# ---- Текстовый подход: что это ----
s = slide_blank(CH_APPR, "Подход 1: текстовый")
textbox(s, 0.49, 1.38, 9.0, 0.28,
        "Агент работает с кодом как с текстом: строки и подстроки, а не функции и вызовы",
        size=13, color=GREY)
labels = [("Задача", ""), ("Поиск", "grep, ripgrep"), ("Чтение файлов", ""), ("Правка", "")]
xs4 = [0.72, 3.01, 5.33, 7.61]
for i, (t, sub) in enumerate(labels):
    card(s, xs4[i], 2.15, 1.67, 0.95, line=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
         text=[(t, {"size": 13.5, "bold": True})] + ([(sub, {"size": 10.5, "color": GREY})] if sub else []))
    if i < 3:
        arrow(s, xs4[i] + 1.8, 2.555)
card(s, 0.49, 3.67, 9.0, 1.37, line=PURPLE, size=12, gap=4, text=[
    ("МЕХАНИЗМЫ ПРАВКИ", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 8}),
    ("•", "Скрипт / Regex по файлам – заменяет подстроку везде, где нашёл"),
    ("•", "Инструмент «Edit_File» – удалить / добавить / заменить; каждое место – отдельный вызов"),
    ("•", "Перезапись файла целиком – модель генерирует новый файл"),
])
anim(s, [0, 1], [2, 3], [4, 5], [6, 7], [8])

# ---- Демо 1: Claude Code ----
s = slide_demo(CH_APPR, "Подход 1: текстовый", DEMO["text"],
               ["grep по объявлениям lint / compileAndLint",
                "скрипт дописывает compile = false старым вызовам (упал, починил через sed) – и заодно в чужой сборке detekt-sample-extensions",
                "двумя правками Edit сливает функции",
                "скрипт rename по подстроке в 131 файле – задевает комментарий и changelog",
                "тесты без флага -Pcompile-test-snippets=true – зелёные",
                "detekt проекта находит длинную строку; откат changelog",
                "отчёт: в списке изменённых – detekt-sample-extensions"])

# ---- Разбор демо 1 ----
s = slide_blank(CH_APPR, "Подход 1: текстовый")
result_table(s, [
    (True, "все 1506 вызовов"),
    (False, "сломал соседний модуль и комментарий"),
], metrics=DEMO["text"]["metrics"])
anim(s, [0, 1], [2, 3], range(4, 8))
notes(s, "Полнота: сильная модель нашла всё, включая два вызова без точки (в 2 из 16 прогонов их пропускали). Корректность: скрипт по подстроке дописал compile = false "
         "и в detekt-sample-extensions – модуль собирается против опубликованной 1.23.7, где параметра нет; rename по подстроке испортил комментарий «not lint for performance reasons». "
         "Сборка зелёная: модуль в основную сборку не входит, тесты шли без флага. Попытки – ошибок по ходу 6, сам исправил 3. Контекст – стартовый, первый запрос.")

# ---- Текстовый подход: плюсы и минусы ----
s = slide_pros_cons("Подход 1: текстовый",
    ["Никаких зависимостей",
     "Скорость",
     "Не требователен к ресурсам",
     "Работает где угодно: CI, worktree, удалённая машина",
     "Любой язык и формат файла"],
    ["Не учитывает семантику языка: перегрузки, переопределения, вызов с точкой и без точки, и т. д.",
     "Применение не проверяет результат: строка заменена – не значит, что поведение не изменилось",
     "Каждое место – отдельная операция, в которой можно ошибиться"], y=1.83, h=2.74)

# ---- Текстовый подход: где ломается ----
s = slide_blank("Подход 1: текстовый", "Полнота и корректность")
textbox(s, 0.49, 1.22, 9.0, 0.26, "Одна строка, но две разные функции", size=12, color=GREY)
card(s, 0.49, 1.7, 9.0, 1.6, line=WHITE, size=12.5, font=MONO, gap=2, margin=0.25, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("detekt-rules-style/…/UnusedParameterSpec.kt", {"color": GREY, "size": 10.5}),
    ("    subject.lint(code)   // Rule.lint – наша функция", {"color": WHITE, "space_after": 6}),
    ("detekt-formatting/…/IndentationSpec.kt", {"color": GREY, "size": 10.5}),
    ("    subject.lint(code)   // FormattingRule.lint – другая функция", {"color": RED}),
])
card(s, 0.49, 3.5, 4.4, 1.6, line=RED, size=12.5, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("Полнота", {"bold": True, "space_after": 4}),
    "Поиск по подстроке: пропускает нужные места и захватывает лишние",
])
card(s, 5.1, 3.5, 4.4, 1.6, line=RED, size=12.5, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("Корректность", {"bold": True, "space_after": 4}),
    "После замены вызов может вести в другую функцию – замена этого не проверяет",
])
anim(s, [0, 1], [2, 3])
notes(s, "Пример из прогонов 2026-10-10: в 4 из 5 агент тронул IndentationSpec.kt – заменил import compileAndLint на import lint; после этого subject.lint(code) в файле "
         "вызывает Rule.lint вместо FormattingRule.lint – текст не изменился, функция другая, сборка зелёная. TODO: подтвердить в IDE (Go to Declaration). "
         "Вызов без точки lint(\"fun f() { $code }\") – вторая ловушка того же рода: в прогоне 29.09 пропущен, в прогонах 10.10 найден. "
         "Вывод: нужен инструмент, который видит функции и их вызовы, а не строки.")

# ---- Сравнение: после текста ----
s = slide_blank("Сравнение", "Сравнение подходов")
compare_table(s, upto="Агент + текстовый подход")
notes(s, "Пауза после подхода. Таблица растёт по ходу доклада: после каждого подхода – колонка.")

slide_subplan(current="LSP-серверы", done={"Текстовый подход"})

# ---- LSP: что это ----
s = slide_blank(CH_APPR, "Подход 2: LSP")
textbox(s, 0.49, 1.29, 9.0, 0.69, [
        "LSP (Language Server Protocol) – протокол между редактором / агентом и сервером языка",
        "Сервер – это отдельный процесс, который разбирает код и знает, к какому символу что относится",
        "Один сервер на один язык",
        ], size=12.5, color=GREY)
labels = [("Клиент", "редактор или агент"), ("Запрос", "«переименуй этот символ»"),
          ("Сервер языка", "резолвит символы"), ("Правка", "список изменений, применяет клиент")]
for i, (t, sub) in enumerate(labels):
    card(s, X4N[i], 2.4, 1.67, 1.0, line=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
         text=[(t, {"size": 14, "bold": True}), (sub, {"size": 10.5, "color": GREY})])
    if i < 3:
        arrow(s, X4N[i] + 1.8, 2.83)
card(s, 0.72, 3.65, 8.56, 1.45, line=GREEN, size=12, gap=4, text=[
    ("Что даёт", {"size": 15, "bold": True, "space_after": 8}),
    ("•", "Переход к определению, поиск использований, переименование по символу"),
    ("•", "Серверы есть для всех основных языков: rust-analyzer, gopls, pyright; Kotlin LSP – пока экспериментальный"),
])
anim(s, [0, 1], [2, 3], [4, 5], [6, 7], [8])

# ---- Демо 2: LSP ----
s = slide_demo(CH_APPR, "Подход 2: LSP", DEMO["lsp"],
               ["объясняет, зачем LSP: grep не отличит lint(KtFile) от lint(String)",
                "rename_symbol через LSP: старый lint → временное имя, «24 changes applied»",
                "сервер заодно удалил чужой импорт – агент заметил по диффу и откатил",
                "detekt-sample-extensions LSP не видит; агент его не трогает",
                "compile = false – смена сигнатуры, в протоколе её нет: скрипт (18 файлов из 116, потом sed)",
                "тесты с флагом; длинная строка → правка → зелёный повтор"])

# ---- LSP: разбор ----
s = slide_blank(CH_APPR, "Подход 2: LSP")
result_table(s, [
    (True, "все 1506 вызовов, по ссылкам"),
    (True, "поведение не изменилось"),
], metrics=DEMO["lsp"]["metrics"])
anim(s, [0, 1], [2, 3], range(4, 8))
notes(s, "Переименование по ссылкам нашло вызовы без точки и не тронуло одноимённые функции. Но правки применяет сервер: он удалил посторонний import assertThat в файле, "
         "где переименовывать было нечего, и не сообщил – агент заметил только по диффу (так в 5 из 5 прогонов с rename_symbol). Из рефакторингов в протоколе только rename: "
         "добавить параметр пришлось скриптом, и скрипт сначала покрыл 18 файлов из 116. По 15 прогонам: 5 не довели задачу до 5/5.")

# ---- LSP: плюсы и минусы (минусы – ограничения протокола) ----
s = slide_pros_cons("Подход 2: LSP",
    ["Нужен только сервер языка",
     "Скорость",
     "Не требователен к ресурсам",
     "Работает где угодно: CI, worktree, удалённая машина",
     "Ошибки компиляции видны сразу после правки"],
    ["Из рефакторингов – только переименование",
     "Строгие рамки протокола LSP",
     "На каждый язык нужен отдельный сервер",
     "Не учитывает межъязыковые связи: например, Kotlin - Java"], y=1.84, h=2.58)
notes(s, "Минусы – уровень протокола. Переименование – единственный запрос рефакторинга, которому клиент передаёт параметры (textDocument/rename); "
         "extract, inline, move бывают только как code actions: сервер сам предлагает список в точке курсора, передать имя или модуль нечем. "
         "Перенос, смена сигнатуры, удаление – в протоколе нет.  Что сервер делает внутри ограничений – своё у каждого: clangd и gopls отказывают, rust-analyzer помечает правку, Kotlin LSP отказал нам. "
         "Вывод: символы видит, оба вызова без точки нашёл; но операция одна, конфликт как повезёт, один язык. Нужен движок.")

# ---- Сравнение: после LSP ----
s = slide_blank("Сравнение", "Сравнение подходов")
compare_table(s, upto="Агент + LSP")
notes(s, "Пауза после подхода. После движка добавится четвёртая колонка.")

slide_subplan(current="Инструменты IDE", done={"Текстовый подход", "LSP-серверы"})

# ---- Движок: что это ----
s = slide_blank(CH_APPR, "Подход 3: движок рефакторинга IDE (IntelliJ)")
card(s, 0.49, 1.43, 9.0, 0.55, line=PURPLE, text="Инструменты IDE позволяют разделить: модель решает «что», движок – «как»",
     size=14, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
textbox(s, 0.49, 2.23, 9.0, 0.26, "Код представлен не текстом, а графом символов: символ – узел, ссылка – ребро к объявлению",
        size=12, color=GREY)
card(s, 0.49, 2.77, 4.4, 1.63, line=RED, size=11, font=MONO, gap=1, text=[
    ("КАК ВИДИТ GREP", {"size": 10.5, "color": RED, "bold": True, "font": FONT, "space_after": 6}),
    ("grep '\\.lint('  → 180 строк:", {"color": LILAC}),
    "  rule.lint(code)      Rule.lint",
    "  subject.lint(code)   FormattingRule.lint",
    "  rule.lint(ktFile)    перегрузка",
    ("  lint(\"fun f() …\")    не найдено", {"color": RED}),
])
card(s, 5.1, 2.77, 4.4, 1.63, line=GREEN, size=11, font=MONO, gap=1, text=[
    ("КАК ВИДИТ IDE", {"size": 10.5, "color": GREEN, "bold": True, "font": FONT, "space_after": 6}),
    ("Rule.lint(content)   99 ссылок", {"color": GREEN}),
    ("Rule.lint(ktFile)    другой символ", {"color": GREY}),
    ("FormattingRule.lint  другой символ", {"color": GREY}),
])
textbox(s, 0.49, 4.66, 9.0, 0.45, [
        "PSI (Program Structure Interface) – модель кода: дерево + ссылки между узлами",
        "IDE держит её в памяти и синхронизирует с редактором"], size=11.5, color=LILAC, align=PP_ALIGN.CENTER)
anim(s, [0, 1], [2], [3, 4])
notes(s, "Агент решает, что переименовать и как поступить с конфликтом; движок ищет ссылки, проверяет конфликт, применяет. Рассматриваем только IntelliJ.")

# ---- Veai ----
s = slide_blank(CH_APPR, "Veai агент в JetBrains IDE")
picture(s, "img/veai-agent.gif", 0.49, 1.27, w=6.05)                 # 800×450 → 6.05×3.4; ролик с veai.ru, 24 с – раскладка Михаила
card(s, 6.83, 1.27, 2.77, 1.9, line=PURPLE, size=10.5, gap=2, margin=0.15, text=[
    ("ИНСТРУМЕНТЫ IDE", {"size": 10, "color": LILAC, "bold": True, "space_after": 4}),
    "Find usages · Rename · Move · Safe delete", "Отладчик", "Инспекции IDE", "…"])
card(s, 6.83, 3.3, 2.77, 1.37, line=GREEN, size=10.5, gap=2, margin=0.15, text=[
    ("ИНСТРУМЕНТЫ АНАЛИЗА", {"size": 10, "color": GREEN, "bold": True, "space_after": 4}),
    "Поиск уязвимостей (SAST)", "Анализ на моргающие тесты", "Создание тестов из исполнения", "…"])
textbox(s, 0.49, 4.85, 9.11, 0.55, "Veai = ИИ + инструменты IDE + инструменты анализа программ", size=14, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
anim(s, [0], [1], [2], [3])
notes(s, "По veai.ru и veai.ru/docs (2026-10-14). Команда – более 50 инженеров, экс-JetBrains и Huawei, экспертиза в статическом и динамическом анализе кода, генерации тестов и символьном исполнении. "
         "Дальше демо движка идёт на Veai.")

# ---- Демо 3: движок ----
s = slide_demo(CH_APPR, "Подход 3: движок рефакторинга IDE", DEMO["engine"],
               ["find_usages на старом lint: 99 мест по ссылкам",
                "Rename lint → lintWithoutCompilation: пробный прогон, 23 файла, 0 конфликтов, применение",
                "Rename compileAndLint → lint: 130 файлов, 0 конфликтов; задвоенные импорты движок схлопнул сам",
                "параметр compile – обычной правкой; compile = false в вызовы – perl по временному имени (рефакторинга «добавить аргумент» нет)",
                "Gradle в терминале падает (JDK 25) → раннер IDE; 18 тестов; apiCheck упал на пустой строке → починил"])

# ---- Разбор демо 3 ----
s = slide_blank(CH_APPR, "Подход 3: движок рефакторинга IDE")
result_table(s, [
    (True, "все 1506 вызовов"),
    (True, "поведение не изменилось"),
], metrics=DEMO["engine"]["metrics"])
anim(s, [0, 1], [2, 3], range(4, 8))
notes(s, "Та же модель, что в демо 1–2. Оба переименования – движком, с пробным прогоном и проверкой конфликтов до записи. detekt-sample-extensions не тронут: "
         "модуля нет в модели проекта IDE, движок его не видит. Ошибки только в сборке и проверке (JDK, пустая строка в API-дампе) – сам рефакторинг не переделывался. "
         "Временное имя от движка сделало текстовую замену точной. Контекст ≈ 20k – оценка по исходникам плагина, в 2–3 раза меньше Claude Code. По 9 прогонам: 9 чисто.")

# ---- Движок: плюсы и минусы ----
s = slide_pros_cons("Подход 3: движок IDE",
    ["Полнота: находит все ссылки, включая другие языки и конфигурации",
     "Корректность: проверяет и показывает список конфликтов до применения",
     "Атомарность: одна операция рефакторинга, один Ctrl + Z"],
    ["Нужна запущенная IDE",
     "Индексация в IDE – не самая быстрая операция",
     "Применение правки – блокировка на запись",
     "Проблематично с Worktree, CI, удалёнными машинами",
     "Есть ограничения: например, Runtime-рефлексия"], y=1.7, h=2.85)
notes(s, "Полнота и корректность – далее по слайду на каждую. Атомарность: либо всё, либо ничего – не бывает «переименовано в 130 файлах из 133»; "
         "если код изменился между поиском и применением, движок не применяет правку и просит повторить поиск («There were changes in code after usages have been found»). "
         "Архитектурные изменения – не рефакторинг в смысле движка; это оговорка, не минус.")

# ---- PSI ----
def dashed_arrow(slide, x1, y1, x2, y2, color=GREEN, width=1.5):
    """Пунктирная стрелка-коннектор (ссылка на схемах)."""
    from pptx.enum.dml import MSO_LINE
    from pptx.oxml.ns import qn as _q
    ln = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    ln.line.dash_style = MSO_LINE.DASH
    ln_el = ln.line._get_or_add_ln()
    ln_el.append(ln_el.makeelement(_q("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return ln


s = slide_blank("Подход 3: движок IDE", "PSI: как IDE видит код")
PX, PY, PW = 0.4, 1.41, 6.2                      # картинка на слайде
SC = PW / 2570                                   # дюймов на пиксель кропа (crop origin 30,260 в screenshot1.png)
def px(x, y):
    return PX + (x - 30) * SC, PY + (y - 260) * SC
picture(s, "img/psi-viewer.png", PX, PY, w=PW)   # 2570×1520 → 6.2×3.67 (раскладка Михаила: 0.4, 1.41)
def outline(x1, y1, x2, y2, col=GREEN):
    ax, ay = px(x1, y1); bx, by = px(x2, y2)
    r = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(ax - 0.04), Inches(ay - 0.04), Inches(bx - ax + 0.08), Inches(by - ay + 0.08))
    r.adjustments[0] = 0.3; r.fill.background(); r.line.color.rgb = col; r.line.width = Pt(2); r.shadow.inherit = False
    return r
outline(585, 1398, 1105, 1462)                   # узел REFERENCE_EXPRESSION
nx, ny = px(1105, 1430)
card(s, 6.92, 2.88, 2.75, 0.72, line=GREEN, size=10.5, font=MONO, gap=1, margin=0.12, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("RuleExtensions.kt", {"size": 9.5, "color": GREY, "font": FONT}),
    ("fun Rule.lint(content, …)", {"color": GREEN, "bold": True}),
])
elbow_arrow(s, [(nx + 0.06, ny), (5.0, ny), (5.0, 3.24), (6.87, 3.24)])
anim(s, [0], [1], [2, 3])
notes(s, "Скриншот PSI Viewer IntelliJ на BracesOnIfStatementsSpec.kt:2213. Дерево – синтаксис, его видит и парсер. Ссылка – то, чего в тексте нет: узел lint знает, к какому объявлению относится; "
         "выбор зависит от типа receiver и импортов. В настоящем дереве у этой строки 15 узлов; индексы хранят, где какое имя встречается, – обход ссылок быстрый.")

# ---- Полнота ----
s = slide_blank("Подход 3: движок IDE", "Полнота")
_i = lambda: len(s.shapes) - 3
def codebox(x, y, w, text, col=WHITE, size=10.5):
    return card(s, x, y, w, 0.5, line=col, size=size, font=MONO, margin=0.08, anchor=MSO_ANCHOR.MIDDLE,
                text=[(text, {"color": col if col is not WHITE else WHITE, "bold": col is GREEN})])
def frame(x, y, w, h, title):
    card(s, x, y, w, h, line=LILAC, width=1.0, size=11, margin=0.15, text=[(title, {"color": LILAC, "bold": True})])
g_k = _i()
frame(0.49, 1.42, 5.8, 3.45, "Kotlin")
codebox(0.65, 1.87, 5.5, "fun Rule.lint(content: String)", GREEN)          # объявление; низ 2.37
g_u = _i()
usages = [(0.65, 1.85, "subject.lint(code)"), (2.65, 1.75, "lint(\"fun f() {}\")"), (4.55, 1.6, "import test.lint")]
for x, w, t in usages:
    codebox(x, 3.77, w, t, size=10)
    elbow_arrow(s, [(x + w / 2, 3.72), (x + w / 2, 2.37)])
g_j = _i()
frame(6.7, 1.42, 2.8, 1.5, "Java")
codebox(6.85, 1.87, 2.5, "lint(rule, code)")
elbow_arrow(s, [(6.8, 2.07), (6.15, 2.07)])
g_x = _i()
frame(6.7, 3.17, 2.8, 1.7, "XML")
codebox(6.85, 3.77, 2.5, "<bean class=\"MyRule\"/>")
elbow_arrow(s, [(6.8, 4.02), (6.5, 4.02), (6.5, 2.27), (6.15, 2.27)], r=0.14)
anim(s, range(g_k, g_u), range(g_u, g_j), range(g_j, g_x), list(range(g_x, _i())))
notes(s, "Объявление одно, ссылки на него – из разных файлов и разных языков: модули Kotlin, Java, XML дают свои узлы и свои ссылки, движок обходит их одинаково. "
         "Java и XML – иллюстрация возможности (в detekt функцию из Java не вызывают; Java-вызов показан со static import). Ссылка – слой поверх дерева: плагин может добавить связь туда, где компилятор её не видит. "
         "Индексы – почему обход быстрый. На некомпилирующемся коде – частично: неразрешённые ссылки помечаются отдельно.")

# ---- Корректность ----
s = slide_blank("Подход 3: движок IDE", "Корректность")
card(s, 0.72, 1.39, 8.56, 0.6, line=PURPLE, text="Правка примеряется на копии в памяти: конфликты видны до того, как изменен файл",
     size=13.5, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
steps = [("Копия вызова", "в памяти, файл не меняется"), ("Переименование", "на копии"),
         ("Проверка", "те же вызовы?"), ("Конфликт или переименование", "")]
for i, (t, sub) in enumerate(steps):
    card(s, X4N[i], 2.29, 1.67, 0.8, line=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
         text=[(t, {"size": 12, "bold": True})] + ([(sub, {"size": 10.5, "color": GREY})] if sub else []))
    if i < 3:
        arrow(s, X4N[i] + 1.8, 2.62)
# пример: по клику первая строка переименовывается и становится неотличима от второй; затем вердикт движка
card(s, 0.72, 3.39, 8.56, 1.0, line=WHITE, size=12, font=MONO, gap=3, margin=0.2, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("rule.compileAndLint(code)   // компилирует сниппет", {"color": WHITE}),
    ("rule.lint(code)             // другая функция, без компиляции", {"color": GREEN}),
])
card(s, 0.72, 3.39, 8.56, 1.0, line=RED, fill=RGBColor(0x1A, 0x12, 0x33), size=12, font=MONO, gap=3, margin=0.2, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("rule.lint(code)             // переименовано", {"color": RED}),
    ("rule.lint(code)             // неразличимы: 78 вызовов меняют смысл", {"color": RED}),
])
card(s, 0.72, 4.64, 8.56, 0.55, line=GREEN, size=12.5, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1,
     text="Движок видит это на копии и останавливается: конфликт")
anim(s, [0], [1], [2, 3], [4, 5], [6, 7], [8], [9], [10])
notes(s, "Механизм для метода: копия выражения-вызова, новое имя, повторный резолв. На кейсе – конфликт объявления.")

# ---- Таблица ----
s = slide_blank("Сравнение", "Сравнение подходов")
compare_table(s)
notes(s, "Не читать вслух. У текста нет символа; у LSP символ есть, но только rename, один язык и конфликты как повезёт; у движка – модель проекта и проверка до правки. "
         "Ниша текста – некомпилирующийся код и нулевая цена входа. Строка контекста – ответ на «сколько стоят описания инструментов».")

# ---- 40 прогонов ----
s = slide_blank("Сравнение", "Сравнение подходов")
RUNS = [("Текст", "Claude Code · 16 прогонов", GREY, [("16 из 16", "решена"), ("5", "чисто"), ("9", "откатов")]),
        ("LSP", "Claude Code + Serena · 15 прогонов", LILAC, [("10 из 15", "решена"), ("7", "чисто"), ("10", "откатов")]),
        ("Инструменты IDE", "Veai · 9 прогонов", GREEN, [("9 из 9", "решена"), ("9", "чисто"), ("0", "откатов")])]
for i, (name, sub, col, nums) in enumerate(RUNS):
    x = X3[i]
    card(s, x, 1.6, 2.85, 3.5, line=col, size=12, margin=0.2, text=[
        (name, {"size": 16, "bold": True, "space_after": 2}),
        (sub, {"size": 10.5, "color": GREY, "space_after": 10}),
    ] + sum([[(v, {"size": 24, "bold": True, "color": col}), (k, {"size": 11, "color": GREY, "space_after": 8})] for v, k in nums], []))
anim(s, [0], [1], [2])
notes(s, "Выборка Максима Паршина: Claude Code 16 прогонов, Claude Code + Serena 15, Veai 9 (5 на Opus 5.5, 4 на GPT-6.1-Sol); два варианта промпта – с тестами и без полного прогона. "
         "«Решена» – 5 из 5 обязательных пунктов чекера. «Чисто» – решена и ни одной лишней правки в итоговом диффе (чужая сборка, документация). "
         "Откаты – агент откатывал собственные правки git checkout, почти всегда после правки чужой сборки detekt-sample-extensions. Veai модуль не видит: его нет в модели проекта IDE. "
         "LSP: 5 прогонов из 15 не довели задачу до 5/5. Время между подходами не сравниваем: промпты и кэш Gradle разные.")

# ---- Переход к Bias ----
s = slide_blank(CH_APPR, "")
picture(s, "img/meme-drake.jpg", 0.49, 1.15, h=4.0)                # 1200×1200 → 4.0×4.0
textbox(s, 4.9, 1.15, 4.6, 4.0, ["Дали агенту классные инструменты, а он ими не пользуется…", "Почему?"], size=22, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Переход: движок показал лучший результат, а терминальные агенты по умолчанию идут в grep и sed. Дальше – почему: инструкции агентов и привычки моделей.")

# ---- Bias агентов (a): Codex и Claude Code ----
s = slide_blank(CH_APPR, "Bias агентов")
source_line(s, 0.49, 1.17, 9.0, "Codex CLI – системный промпт")
picture(s, "img/codex-prompt.png", 0.49, 1.59, w=5.9)                # 1800×464 → 5.9×1.52
source_line(s, 0.49, 3.39, 9.0, "Claude Code – системный промпт")
picture(s, "img/claude-code-bypass-issue.png", 0.49, 3.76, w=5.9)    # 1640×329 → 5.9×1.18
textbox(s, 6.6, 1.97, 2.9, 0.77, [("«Для поиска используй rg, для редактирования – скрипты»", {"size": 14, "bold": True, "color": LILAC})],
        size=14, anchor=MSO_ANCHOR.MIDDLE)
textbox(s, 6.6, 3.85, 2.9, 1.0, [("«Работай через bash: для чтения – cat, для поиска – grep, для редактирования – sed»", {"size": 14, "bold": True, "color": LILAC})],
        size=14, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Это инструкция агента, а не выбор модели: та же модель в другом агенте инструменты вызывает. Единственный инструмент правки у терминальных агентов – замена строки или патч. "
         "Источники: github.com/openai/codex, codex-rs/core/gpt_5_codex_prompt.md (тот же текст в промптах для gpt-5.2); блок bypass/auto – по issue anthropics/claude-code #88475 (версии 2.1.228+). "
         "Документация Claude Code: на macOS и Linux Glob и Grep убраны из набора по умолчанию – поиск через find и grep в Bash. OpenCode: описание инструмента Edit предлагает replaceAll, «if you want to rename a variable».")

# ---- Bias агентов (b): Cursor ----
s = slide_blank(CH_APPR, "Bias агентов")
source_line(s, 0.49, 1.41, 9.0, "Форум Cursor, июль 2026 – ответы сотрудников Cursor")
picture(s, "img/cursor-forum-colin.png", 0.49, 1.97, w=6.1)          # 2340×534 → 6.1×1.39
picture(s, "img/cursor-forum-maximep.png", 0.49, 3.7, w=6.1)         # 2340×382 → 6.1×1.0
textbox(s, 6.8, 1.97, 2.7, 1.39, [
    ("«Семантический индекс (RAG) отключён в пользу grep»", {"size": 14, "bold": True, "color": LILAC}),
], size=14, anchor=MSO_ANCHOR.MIDDLE)
textbox(s, 6.8, 3.7, 2.7, 1.0, [
    ("«Модели научились искать с помощью grep, индекс (RAG) не нужен»", {"size": 14, "bold": True, "color": LILAC}),
], size=14, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Не промпт, а продуктовое решение: Cursor убрал семантический индекс, потому что модели справляются grep'ом. forum.cursor.com/t/…/165899")

# ---- Bias агентов (c): Serena ----
s = slide_blank(CH_APPR, "Bias агентов")
textbox(s, 0.49, 1.15, 9.0, 0.25, "Документация Serena – MCP-сервера с инструментами LSP для агентов, октябрь 2026", size=11, color=GREY)
picture(s, "img/serena-docs-attention.png", 0.49, 1.52, h=3.75)      # 1640×1082 → 5.68×3.75
textbox(s, 6.4, 2.59, 3.1, 1.61, [
    ("«Почти невозможно убедить Opus 4.7 пользоваться Serena»", {"size": 16, "bold": True, "color": LILAC, "space_after": 10}),
    ("Авторы Serena дают отдельный системный промпт, чтобы это исправить", {"size": 12, "color": GREY}),
], size=16, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Переход к LSP: Serena даёт агенту LSP, и её авторы пишут, что модель к ней сама не идёт. Демо LSP – на Claude Code + Serena. oraios.github.io/serena/02-usage/030_clients.html")

# ---- Бенчмарк по моделям ----
s = slide_blank(CH_APPR, "Bias моделей")
def mini_table(x, y, cols, rows, label_w, col_w):
    """Мини-таблица: строки «Поиск в коде / Рефакторинг», в ячейках – доля прогонов с инструментами IDE."""
    for i, c in enumerate(cols):
        textbox(s, x + label_w + i * col_w, y, col_w, 0.3, c, size=10, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
    hline(s, x, y + 0.32, label_w + col_w * len(cols), color=GREY, width=0.5)
    for r, (name, vals) in enumerate(rows):
        yy = y + 0.38 + r * 0.42
        textbox(s, x, yy, label_w, 0.42, name, size=11, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
        for i, v in enumerate(vals):
            textbox(s, x + label_w + i * col_w, yy, col_w, 0.42, v, size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
card(s, 0.49, 1.4, 4.4, 3.7, line=WHITE, size=12, gap=8, text=[
    ("Наш эксперимент, октябрь 2026", {"size": 13, "bold": True, "color": LILAC, "space_after": 4}),
    ("9 задач рефакторинга на Java/Kotlin", {"color": GREY}),
    ("15 прогонов на модель", {"color": GREY, "space_after": 8}),
])
mini_table(0.7, 2.75, ["Sol 6.1", "Opus 5.5", "GLM 5.3"],
           [("Поиск в коде", ["53 %", "0 %", "13 %"]), ("Рефакторинг", ["67 %", "47 %", "40 %"])], 1.6, 0.8)
textbox(s, 0.7, 4.3, 4.0, 0.6, "Доля прогонов, где агент пользовался инструментами IDE", size=10.5, color=GREY)
card(s, 5.1, 1.4, 4.4, 3.7, line=WHITE, size=12, gap=8, text=[
    ("Эксперимент AgentConnect, август 2026", {"size": 13, "bold": True, "color": LILAC, "space_after": 4}),
    ("Моделям Claude дали инструменты LSP рядом с grep", {"color": GREY, "space_after": 8}),
])
mini_table(5.3, 2.75, ["Claude"], [("Поиск в коде", ["0–6 %"]), ("Рефакторинг", ["3 %"])], 2.2, 1.0)
textbox(s, 5.3, 4.3, 4.0, 0.6, [("Доля задач, где агент выбрал LSP. Если LSP навязать – успешных задач 89 % вместо 100 %", {}),
                                ("agentconnect.md/blog/grep-beat-lsp-harness", {"size": 9.5})], size=10.5, color=GREY)
anim(s, range(0, 14), range(14, 22))
notes(s, "Бенчмарк Максима Паршина (colleague-package 2026-10-16): 54 прогона, сопоставимая выборка 15 на модель; reasoning medium; skills и делегирование выключены. "
         "Общего избегания инструментов нет: move на Kafka и rename на Gson/Jackson применяют все. Разница – в стиле: Sol строит IDE-first (usages, структура, run configurations), "
         "Opus и GLM исследуют и проверяют через shell. Оговорки: выборка небольшая и неоднородная, описания инструментов сами подталкивают к IDE. "
         "AgentConnect – оговорки: пилот, 2–3 прогона на ячейку, rename через LSP не тестировали. Тезис статьи – «agent capability = model × harness»: "
         "модель сама к LSP не идёт, а если заставить без доработки агента – становится хуже.")

slide_plan(done={CH_INTRO, CH_APPR}, current="Итоги")

# ---- Что выбрать ----
s = slide_blank("Итоги", "Что выбрать")
rows = [
    ("Текст", "Нет IDE: CI, скрипты, конфиги, чужая машина. Одно имя – одна функция, мест немного", GREY),
    ("LSP", "Один язык, есть сервер. Нужны все использования и переименование; остальное – текстом", LILAC),
    ("Движок IDE", "Есть IDE и индексы. Несколько языков, одноимённые функции, сотни мест – нужны гарантии", GREEN),
]
y = 1.4
for name, text, col in rows:
    card(s, 0.49, y, 2.2, 1.05, line=col, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, text=[(name, {"size": 16, "bold": True})])
    card(s, 2.85, y, 6.65, 1.05, line=col, anchor=MSO_ANCHOR.MIDDLE, size=12.5, text=text)
    y += 1.2
anim(s, [0, 1], [2, 3], [4, 5])
notes(s, "Итог по выбору подхода – просьба оргкомитета. Ни один не «правильный»: у каждого ситуация, где он уместен.")

# ---- Veai CLI ----
s = slide_blank("Итоги", "Veai CLI")
picture(s, "img/veai-cli.gif", 0.49, 1.3, w=6.0)                     # 800×500 → 6.0×3.75; ролик с veai.ru/cli, 29 с
card(s, 6.87, 1.65, 2.8, 3.04, line=PURPLE, size=11.5, gap=8, margin=0.2, text=[
    ("«Русский Claude Code»", {"size": 14, "bold": True, "space_after": 8}),
    ("•", "Все топовые модели в одной подписке"),
    ("•", "Работает без VPN"),
    ("•", "Оплата русской картой"),
    ("•", "Джентльменский набор: субагенты, правила, скиллы, MCP, …"),
    ("•", "Управление из Telegram"),
    ("veai.ru/cli", {"size": 14, "bold": True, "color": LILAC, "space_before": 6}),
])
notes(s, "Словами: текстовый подход неотъемлем – много задач, где семантика не нужна (конфиги, скрипты, документация, CI без IDE), поэтому у нас есть CLI. "
         "По veai.ru/cli (2026-10-16): модели Anthropic, OpenAI, Gemini, GLM, Kimi, Qwen, DeepSeek; Telegram-управление; автоподбор модели; прогрев кэша.")

# ---- Советы ----
s = slide_blank("Итоги", "Советы пользователям агентов")
card(s, 0.49, 1.7, 4.4, 2.2, line=GREEN, size=12.5, gap=10, text=[
    ("Рефакторинг – отдельным коммитом", {"size": 15, "bold": True}),
    "Отдельно от изменения поведения. Ревью читаемо, откат возможен",
])
card(s, 5.1, 1.7, 4.4, 2.2, line=GREEN, size=12.5, gap=10, text=[
    ("Автоматическая проверка – обязательна", {"size": 15, "bold": True}),
    "Тесты, компиляция, линтер. В кейсе сборка была зелёной при изменённом поведении – хелперы тестами не покрыты",
])
anim(s, [0], [1])

# ---- Summary ----
s = slide_blank("Итоги", "Главные мысли")
three_cards(s, ["Рефакторинг – преобразование с гарантиями", "Дело в инструментах, не в модели", "Инструмент рефакторинга уже есть – в IDE"], [
    "Полнота при поиске, корректность при применении",
    "Та же модель с grep и с движком – разный результат и разная цена. Считайте ходы, время и токены, а не только зелёные тесты",
    "Полнота, конфликты до правки, атомарность. Модель решает «что», движок делает «как». Отдайте его агенту",
], y=1.45, h=3.0)
anim(s, [0], [1], [2])

# ---- Спасибо ----
s = prs.slides.add_slide(L_TEXT)
_clear_placeholders(s)
bg_glow(s)
_chrome(s, TALK_SHORT)
textbox(s, 0.4, 1.6, 6.9, 2.0, ["Спасибо", "за внимание"], size=48, color=WHITE)
textbox(s, 0.4, 4.0, 6.0, 1.0, [("Вопросы?", {"size": 20, "color": LILAC, "bold": True})], size=20)
textbox(s, 5.17, 4.39, 4.42, 0.82, [
    ("Попробовать агента: veai.ru", {"bold": True}),
    ("Telegram-канал: @veai_devs", {"bold": True}),
], size=13, color=WHITE)

# ---- Приложение: MCP (после «Спасибо», в план не входит) ----
s = slide_blank("Готовые решения", "MCP-серверы для рефакторинга")
textbox(s, 0.49, 1.49, 9.0, 0.46, [
        "MCP (Model Context Protocol) – протокол общения LLM с инструментами",
        "MCP-сервер – сервер, отвечающий на запросы LLM по MCP"],
        size=12, color=GREY)
mcp_cols = [("MCP-СЕРВЕР INTELLIJ", "Встроен в IDE с версии 2025.2, включается в настройках",
             [("Rename", "переименование символа", GREEN), ("Move", "нет", RED), ("Safe Delete", "нет", RED)]),
            ("MCP-СЕРВЕР VEAI", "Встроен в агента с версии Veai 5.17, бесплатно",
             [("Rename", "переименование символа", GREEN), ("Move", "перенос класса, функции или пакета с обновлением ссылок", GREEN), ("Safe Delete", "удаление с проверкой использований", GREEN)])]
for i, (t, ver, tools) in enumerate(mcp_cols):
    paras = [(t, {"size": 10.5, "color": LILAC, "bold": True, "space_after": 6}), (ver, {"color": GREY, "space_after": 10})]
    for name, what, col in tools:
        paras.append((f"{name} – {what}", {"color": col if col is RED else WHITE, "space_after": 6}))
    card(s, [0.49, 5.1][i], 2.66, 4.4, 2.13, line=WHITE, size=12, gap=4, text=paras)
anim(s, [0], [1], [2])


# ---------- анимация: только там, где по клику появляется аргумент ----------
for _sl in prs.slides:
    _spec = ANIM.get(_sl.slide_id)
    if _spec:
        _content = list(_sl.shapes)[3:]
        animate(_sl, [[_content[i] for i in g] for g in _spec])

# ---- Запасной вариант перехода к Bias (мем «distracted boyfriend») ----
s = slide_blank("Запасной слайд", "")
picture(s, "img/meme-distracted.jpg", 0.49, 1.15, h=3.2)           # 1200×800 → 4.8×3.2
textbox(s, 5.5, 1.15, 4.0, 3.2, ["Дали агенту классные инструменты, а он ими не пользуется…", "Почему?"], size=22, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Запасной вариант слайда 39: если «Drake» не зайдёт, поставить этот.")

# ---------- язык текста: без lang="ru" PowerPoint проверяет орфографию как английскую ----------
from pptx.oxml.ns import qn as _qn
for _sl in prs.slides:
    for _r in _sl._element.iter(_qn("a:rPr")):
        _r.set("lang", "ru-RU")
    for _e in _sl._element.iter(_qn("a:endParaRPr")):
        _e.set("lang", "ru-RU")

prs.save(OUT)
print("слайдов:", len(prs.slides), "→", OUT)
