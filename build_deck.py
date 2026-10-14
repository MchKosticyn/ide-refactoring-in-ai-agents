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
OUT = "Доклад_ Рефакторинг в ИИ-агентах.pptx"

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


def slide_demo(kicker, title, steps, metrics=None, sub=None, actor="ДЕЙСТВИЯ", meta=None):
    """Демо: слева область под видео, справа – метрики и шаги."""
    s = slide_blank(kicker, title)
    paras = [("TODO: видео", {"size": 12, "color": AMBER, "bold": True, "space_after": 8})]
    if sub:
        paras.append((sub, {"size": 12, "color": GREY}))
    if meta:
        card(s, 0.49, 1.35, 5.9, 3.2, line=AMBER, text=paras)
        lines = [ln.replace("\n", " ") for ln in meta.split("\n")]
        card(s, 0.49, 4.65, 5.9, 0.5, line=LILAC, size=11, gap=0, margin=0.15, anchor=MSO_ANCHOR.MIDDLE,
             text=[(" · ".join(lines), {"color": todo_color(meta, WHITE)})])
    else:
        card(s, 0.49, 1.35, 5.9, 3.8, line=AMBER, text=paras)
    y = 1.35
    if metrics:
        mh = 0.42 + 0.27 * len(metrics)
        card(s, 6.6, y, 2.9, mh, line=PURPLE, size=11, gap=2, margin=0.15,
             text=[("МЕТРИКИ", {"size": 10, "color": LILAC, "bold": True, "space_after": 4})]
                  + [(f"{k}: {v}", {"color": todo_color(v)}) for k, v in metrics])
        y += mh + 0.1
    card(s, 6.6, y, 2.9, 5.15 - y, line=WHITE, size=11, gap=5, margin=0.15,
         text=[(actor, {"size": 10, "color": LILAC, "bold": True, "space_after": 6})]
              + [(f"{i + 1}.  {st}", {}) for i, st in enumerate(steps)])
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
        w = 9.01 / n
        for i, (k, v) in enumerate(metrics):
            card(s, 0.49 + i * w, 4.0, w - 0.1, 1.0, line=PURPLE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
                 text=[(k, {"size": 10, "color": LILAC}), (v, {"size": 16, "bold": True, "color": todo_color(v)})])
    if banner:
        bottom_banner(s, banner, y=4.75, size=13)


def slide_pros_cons(kicker, pros, cons, note=None):
    s = slide_blank(kicker, "Плюсы и минусы")
    long = max(len(p) for p in pros + cons) > 80 or len(pros) + len(cons) > 6
    size, gap = (12, 7) if long else (13.5, 9)
    card(s, 0.49, 1.4, 4.4, 3.7, line=GREEN, size=size, gap=gap,
         text=[("Плюсы", {"size": 17, "bold": True, "space_after": 12})] + [("•", p) for p in pros])
    card(s, 5.1, 1.4, 4.4, 3.7, line=RED, size=size, gap=gap,
         text=[("Минусы", {"size": 17, "bold": True, "space_after": 12})] + [("•", c) for c in cons])
    anim(s, [0], [1])
    if note:
        notes(s, note)
    return s


PLAN = ["Введение", "Подходы к рефакторингу в агентах", "MCP-серверы для рефакторинга", "Итоги"]
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
        if i < 2 and APPROACHES[i + 1] == current:
            arrow(s, x + w + 0.2, 2.58, w=0.22, h=0.14, color=GREY)
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
    ("Полнота", ["✓", "✗ 2 из 78", "✓", "TODO"]),
    ("Корректность", ["✓", "✗ двойник, комментарий", "TODO", "TODO"]),
    ("Время", ["≈ 30 мин", "13,7 мин", "TODO", "TODO"]),
    ("Стоимость", ["≈ $30", "$4.41", "TODO", "TODO"]),
    ("Ходы", ["–", "53", "TODO", "TODO"]),
    ("Токены", ["–", "3,9M", "TODO", "TODO"]),
    ("Стартовый контекст", ["–", "TODO", "TODO", "TODO"]),
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
M_TODO = [("Время", "TODO"), ("Стоимость", "TODO"), ("Ходы", "TODO"), ("Токены", "TODO"), ("Стартовый контекст", "TODO")]

# ---- Титул ----
s = slide_title(["Использование инструментов", "рефакторинга IDE в ИИ-агентах"],
                ["Спикер: Михаил Костицын", "Эксперт: Максим Паршин"])

# ---- О нас ----
s = slide_blank(TALK_SHORT, "О нас")
people = [("Михаил Костицын", "Спикер", ["Ведущий разработчик Veai – ИИ-агента для разработчиков",
                                        "Статический анализ и формальная верификация; ИИ-агенты с 2023 года"]),
          ("Максим Паршин", "Эксперт доклада", ["TODO", "TODO"])]
for i, (name, role, lines) in enumerate(people):
    x = 0.49 + i * 4.6
    card(s, x, 1.45, 4.4, 3.55, line=PURPLE if i == 0 else WHITE)
    card(s, x + 0.25, 1.7, 1.5, 1.5, line=AMBER, size=10, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.05,
         text=[("TODO: фото", {"color": AMBER, "bold": True})])
    textbox(s, x + 1.95, 1.75, 2.3, 0.9, [(name, {"size": 16, "bold": True}), (role, {"size": 12, "color": LILAC})], size=16)
    textbox(s, x + 0.25, 3.4, 3.9, 1.5, [(ln, {"color": todo_color(ln, WHITE)}) for ln in lines], size=12.5, color=WHITE, space_after=6)

# ---- план ----
slide_plan(current=CH_INTRO)

# ---- Рамка ----
s = slide_blank(CH_INTRO, "Определение")
card(s, 0.72, 1.71, 8.56, 0.81, line=PURPLE, text="Рефакторинг – это изменение структуры кода без изменения поведения программы",
     size=15, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
card(s, 0.72, 2.82, 4.1, 1.39, line=WHITE, size=12.5, gap=8, text=[
    ("Задача 1: найти все места", {"size": 15, "bold": True, "space_after": 10}),
    "Нужна полнота: 97 % не годится",
    ("78 вызовов × 97 % = 2 пропуска", {"color": LILAC, "bold": True}),
])
card(s, 5.18, 2.82, 4.1, 1.39, line=WHITE, size=12.5, gap=8, text=[
    ("Задача 2: применить", {"size": 15, "bold": True, "space_after": 10}),
    "Нужна корректность: правка не меняет поведение программы",
])
anim(s, [0], [1], [2])
notes(s, "Каждый проход дальше разбираем одинаково: найдены ли все места, только ли нужные правки, сохранено ли поведение – и цена: время, стоимость, ходы, токены, стартовый контекст.")

# ---- Зачем это QA ----
s = slide_blank(CH_INTRO, "Зачем это QA")
card(s, 0.49, 1.4, 9.0, 2.45, line=WHITE, size=16, gap=12, margin=0.35, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("•", "Миграция фреймворка: Selenium → Playwright, JUnit 4 → 5"),
    ("•", "Наведение структуры: PageObject, композиция фикстур"),
    ("•", "Смена тестового API проекта: хелперы, базовые классы"),
    ("•", "…"),
])
card(s, 0.49, 4.1, 9.0, 0.9, line=RED, size=15, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, text=
     "Зелёные тесты не всегда доказывают, что поведение то же")
anim(s, [0], [1])

# ---- Кейс ----
s = slide_blank(CH_INTRO, "Кейс: detekt")
textbox(s, 0.49, 1.29, 9.0, 0.27, "detekt – статический анализатор Kotlin", size=12.5, color=GREY)
card(s, 0.49, 1.95, 9.0, 1.15, line=WHITE, size=11.5, font=MONO, gap=1, margin=0.15, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("fun Rule.compileAndLint(code)", {"color": LILAC, "bold": True}),
    ("    сначала компилирует сниппет компилятором Kotlin (только в CI), потом прогоняет правило", {"color": GREY, "font": FONT}),
    ("fun Rule.lint(code)", {"color": LILAC, "bold": True}),
    ("    сразу прогоняет правило: сниппет может быть неполным и не компилироваться", {"color": GREY, "font": FONT}),
])
card(s, 0.49, 3.5, 9.0, 1.57, line=PURPLE, size=12, gap=3, text=[
    ("ЗАДАЧА: PR #7873", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 6}),
    ("Оставить одну функцию: lint(code, compile = true)", {"bold": True}),
    "compileAndLint(code) → lint(code) – 1428 вызовов",
    "старый lint(code) → lint(code, compile = false) – 78 вызовов, чтобы тесты работали как раньше",
    ("Не трогать: перегрузку lint(ktFile) – 36 вызовов, FormattingRule.lint – 66 вызовов", {"color": GREY}),
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
s = slide_blank(CH_INTRO, "Инструменты рефакторинга IDE")
tools = [("Rename", "⇧F6 – переименовать символ и все его использования"),
         ("Change Signature", "⌘F6 – добавить параметр, подставить значение в вызовы"),
         ("Safe Delete", "⌘⌫ – удалить, если использований нет; иначе показать их")]
for i, (t, sub) in enumerate(tools):
    x = X3[i]
    card(s, x, 1.25, 2.85, 2.4, line=AMBER, size=11, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER,
         text=[("TODO: gif", {"color": AMBER, "bold": True})])
    card(s, x, 3.8, 2.85, 1.3, line=WHITE, size=11, gap=3, margin=0.15, anchor=MSO_ANCHOR.MIDDLE,
         text=[(t, {"size": 14, "bold": True, "space_after": 3}), (sub, {"color": GREY})])
anim(s, [0, 1], [2, 3], [4, 5])
notes(s, "Три инструмента, которые понадобятся в кейсе. Официальные gif JetBrains (блог 2020, светлая тема, Java) есть только для Rename и Change Signature – лучше записать свои на detekt в тёмной теме.")

# ---- Человек в IDE: разбор ----
s = slide_blank(CH_INTRO, "Подход 0: человек с IDE")
result_table(s, [
    (True, ""),
    (True, ""),
], metrics=[("Время", "≈ 30 мин"), ("Стоимость", "≈ $30")])
textbox(s, 0.49, 5.1, 9.0, 0.35, "Можно ли быстрее?", size=14, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
anim(s, [0, 1], [2, 3], [4, 5], [6])
notes(s, "Rename в занятое имя – IDE останавливает и показывает конфликт. Это то, что будем искать у агентов. "
         "Время и стоимость – по ускоренной записи с таймером. Порядок величины подтверждает история PR: оба коммита автора сделаны за 20 минут 3 января 2025 (11:12 и 11:32); "
         "при ставке $60/ч это около $30.")

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
slide_subplan(current="Текстовый подход")

# ---- Текстовый подход: что это ----
s = slide_blank(CH_APPR, "Подход 1: текстовый")
textbox(s, 0.49, 1.38, 9.0, 0.28,
        "Агент работает с кодом как с текстом: строки и подстроки, а не функции и вызовы",
        size=13, color=GREY)
labels = [("Задача", "текстом"), ("Поиск", "по тексту: grep, ripgrep"),
          ("Чтение", "найденных файлов"), ("Правка", "текстом"), ("Результат", "diff")]
xs5 = [0.49, 2.35, 4.21, 6.07, 7.93]
for i, (t, sub) in enumerate(labels):
    card(s, xs5[i], 2.15, 1.58, 0.95, line=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
         text=[(t, {"size": 13.5, "bold": True}), (sub, {"size": 10.5, "color": GREY})])
    if i < 4:
        arrow(s, xs5[i] + 1.62, 2.55, w=0.2, h=0.12)
card(s, 0.49, 3.67, 9.0, 1.37, line=PURPLE, size=12, gap=4, text=[
    ("МЕХАНИЗМЫ ПРАВКИ", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 8}),
    ("•", "Скрипт / regex по файлам – заменяет подстроку везде, где нашёл"),
    ("•", "Инструмент «правка фрагмента» – удалить / добавить / заменить; каждое место – отдельный вызов"),
    ("•", "Перезапись файла целиком – модель генерирует новый файл"),
])
anim(s, range(0, 10), [10])

# ---- Демо 1: Claude Code ----
s = slide_demo(CH_APPR, "Подход 1: текстовый",
               ["grep по имени функции",
                "Чтение файлов",
                "sed / скрипт по файлам",
                "Точечные правки",
                "Отчёт"],
               metrics=[("Время", "13,7 мин"), ("Стоимость", "$4.41"), ("Ходы", "53"),
                        ("Токены", "3,9M"), ("Стартовый контекст", "TODO")],
               sub=None, meta="Агент: Claude Code\nМодель: Claude Opus (TODO версия)")

# ---- Разбор демо 1 ----
s = slide_blank(CH_APPR, "Подход 1: текстовый")
result_table(s, [
    (False, "2 из 78 пропущены"),
    (False, "двойник и комментарий; сборка зелёная"),
], metrics=[("Время", "13,7 мин"), ("Стоимость", "$4.41"), ("Ходы", "53"), ("Токены", "3,9M"), ("Контекст", "TODO")])
anim(s, [0, 1], [2, 3], range(4, 9))
notes(s, "Агент правильно выбрал порядок: сначала старые lint, потом rename; двойник в другом модуле обошёл. Промах – там, где имя не выглядит как вызов метода. "
         "Инструменты агента – поиск по тексту и правка текста.")

# ---- Текстовый подход: плюсы и минусы ----
s = slide_pros_cons("Подход 1: текстовый",
    ["Ничего не нужно: ни индекса, ни IDE, ни запущенного проекта",
     "Работает где угодно: CI, worktree, удалённая машина",
     "Любой язык и формат файла"],
    ["Одно имя – разные функции: FormattingRule.lint и комментарий попадают в замену",
     "Та же функция – другая запись: вызов без точки в замену не попадает",
     "Замена не проверяет результат: строка заменена – не значит, что вызов ведёт туда же",
     "Каждое место – отдельная операция"])

# ---- Текстовый подход: где ломается ----
s = slide_blank("Подход 1: текстовый", "Где ломается")
textbox(s, 0.49, 1.22, 9.0, 0.26, "Одна строка, но две разные функции", size=12, color=GREY)
card(s, 0.49, 1.7, 9.0, 1.6, line=WHITE, size=12.5, font=MONO, gap=2, margin=0.25, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("detekt-rules-style/…/UnusedParameterSpec.kt", {"color": GREY, "size": 10.5}),
    ("    subject.lint(code)   // Rule.lint – та, что меняем", {"color": WHITE, "space_after": 6}),
    ("detekt-formatting/…/IndentationSpec.kt", {"color": GREY, "size": 10.5}),
    ("    subject.lint(code)   // FormattingRule.lint – другая функция", {"color": RED}),
])
card(s, 0.49, 3.5, 4.4, 1.6, line=RED, size=12.5, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("Полнота", {"bold": True, "space_after": 4}),
    "Поиск по подстроке: пропускает одни места и захватывает чужие",
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
s = slide_blank("Сравнение", "Результат и цена")
compare_table(s, upto="Агент + текстовый подход")
notes(s, "Пауза после подхода. Таблица растёт по ходу доклада: после каждого подхода – колонка.")

slide_subplan(current="LSP-серверы", done={"Текстовый подход"})

# ---- LSP: что это ----
s = slide_blank(CH_APPR, "Подход 2: LSP")
textbox(s, 0.49, 1.25, 9.0, 0.9, [
        "LSP (Language Server Protocol) – протокол между редактором или агентом и сервером языка",
        "Сервер – отдельный процесс, который разбирает код и знает, какой текст к какому символу относится",
        "Один сервер на язык",
        ], size=12.5, color=GREY)
labels = [("Клиент", "редактор или агент"), ("Запрос", "«переименуй символ здесь»"),
          ("Сервер языка", "резолвит символы"), ("Правка", "список текстовых изменений; применяет клиент")]
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
anim(s, range(0, 8), [8])

# ---- Демо 2: LSP ----
s = slide_demo(CH_APPR, "Подход 2: LSP",
               ["Запрос references по символу",
                "Список мест от сервера",
                "Правки текстом по списку",
                "Rename по протоколу – если использует"],
               metrics=M_TODO, meta="Агент: Claude Code + Serena (инструменты LSP)\nМодель: TODO")
notes(s, "Перед видео – одна фраза: LSP – сервер, который резолвит символы языка; агент спрашивает его, где символ используется, и просит переименовать.")

# ---- LSP: разбор ----
s = slide_blank(CH_APPR, "Подход 2: LSP")
result_table(s, [
    (True, "99 ссылок, оба без точки"),
    (None, "TODO по записи"),
], metrics=[(k, v) for k, v in M_TODO])
anim(s, [0, 1], [2, 3], range(4, 9))
notes(s, "Rename отказал в нашем прогоне – лучше, чем применить без предупреждения. Но это выбор сервера, а не правило протокола, и продолжить после отказа нечем. "
         "Из пилота AgentConnect (август 2026): LSP точнее grep, но полнота у обоих около 0,66 – LSP не гарантия, что найдено всё.")

# ---- LSP: плюсы и минусы (минусы – ограничения протокола) ----
s = slide_pros_cons("Подход 2: LSP",
    ["Легковесность: нужен только сервер языка",
     "Поиск по символу: находит вызов без точки, различает одноимённые",
     "Ошибки компиляции видны сразу после правки"],
    ["Из рефакторингов – только переименование",
     "Сервер не знает о правках, пока клиент не сообщит",
     "Один сервер – один язык",
     "Конфликт – как решит сервер; «продолжить» нет"])
notes(s, "Минусы – уровень протокола. Переименование – единственный запрос рефакторинга, которому клиент передаёт параметры (textDocument/rename); "
         "extract, inline, move бывают только как code actions: сервер сам предлагает список в точке курсора, передать имя или модуль нечем. "
         "Перенос, смена сигнатуры, удаление – в протоколе нет.  Что сервер делает внутри ограничений – своё у каждого: clangd и gopls отказывают, rust-analyzer помечает правку, Kotlin LSP отказал нам. "
         "Вывод: символы видит, оба вызова без точки нашёл; но операция одна, конфликт как повезёт, один язык. Нужен движок.")

# ---- Сравнение: после LSP ----
s = slide_blank("Сравнение", "Результат и цена")
compare_table(s, upto="Агент + LSP")
notes(s, "Пауза после подхода. После движка добавится четвёртая колонка.")

slide_subplan(current="Инструменты IDE", done={"Текстовый подход", "LSP-серверы"})

# ---- Движок: что это ----
s = slide_blank(CH_APPR, "Подход 3: движок рефакторинга IDE (IntelliJ)")
card(s, 0.49, 1.25, 9.0, 0.55, line=PURPLE, text="Инструменты IDE позволяют разделить: модель решает «что», движок – «как»",
     size=14, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
textbox(s, 0.49, 2.0, 9.0, 0.5,
        "Код представлен не текстом, а графом символов: символ – узел, ссылка – ребро к объявлению. «Все использования» – обход ссылок",
        size=12, color=GREY)
card(s, 0.49, 2.77, 4.4, 1.63, line=RED, size=11, font=MONO, gap=1, text=[
    ("КАК ВИДИТ GREP", {"size": 10.5, "color": RED, "bold": True, "font": FONT, "space_after": 6}),
    ("grep '\\.lint('  → 180 строк:", {"color": LILAC}),
    "  rule.lint(code)      Rule.lint",
    "  subject.lint(code)   FormattingRule.lint",
    "  rule.lint(ktFile)    перегрузка",
    ("  lint(\"fun f() …\")   – не видно", {"color": RED}),
])
card(s, 5.1, 2.77, 4.4, 1.63, line=GREEN, size=11, font=MONO, gap=1, text=[
    ("КАК ВИДИТ IDE", {"size": 10.5, "color": GREEN, "bold": True, "font": FONT, "space_after": 6}),
    ("Rule.lint(content)   99 ссылок", {"color": GREEN}),
    ("  в т. ч. 2 без точки", {"color": GREEN}),
    ("Rule.lint(ktFile)    другой символ", {"color": GREY}),
    ("FormattingRule.lint  другой символ", {"color": GREY}),
    ("Rule.compileAndLint  1428 ссылок", {"color": GREEN}),
])
textbox(s, 0.49, 4.66, 9.0, 0.4,
        "PSI (Program Structure Interface) – модель кода: дерево + ссылки между узлами. IDE держит её в памяти и синхронизирует с редактором",
        size=11.5, color=LILAC, align=PP_ALIGN.CENTER)
anim(s, [0, 1], [2], [3, 4])
notes(s, "Агент решает, что переименовать и как поступить с конфликтом; движок ищет ссылки, проверяет конфликт, применяет. Рассматриваем только IntelliJ.")

# ---- Veai ----
s = slide_blank(CH_APPR, "Veai агент")
card(s, 0.49, 1.3, 9.0, 1.1, line=PURPLE, size=13, gap=6, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("•", "ИИ-агент для разработчиков, работает внутри IDE JetBrains"),
    ("•", "Инструменты – те же, что у IDE: поиск по символам, рефакторинги, диагностика"),
])
card(s, 0.49, 2.55, 9.0, 2.55, line=AMBER, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER,
     text=[("TODO: скриншот Veai внутри IDE с открытым кейсом detekt", {"size": 14, "color": AMBER, "bold": True})])

# ---- Демо 3: движок ----
s = slide_demo(CH_APPR, "Подход 3: движок рефакторинга IDE",
               ["Change Signature: параметр compile, false в вызовах",
                "Rename: конфликт «уже объявлена»",
                "Решение по конфликту",
                "Применение"],
               metrics=M_TODO, meta="Агент: Veai\nМодель: TODO")

# ---- Разбор демо 3 ----
s = slide_blank(CH_APPR, "Подход 3: движок рефакторинга IDE")
result_table(s, [
    (None, "TODO по записи"),
    (None, "TODO по записи"),
], metrics=[(k, v) for k, v in M_TODO])
anim(s, [0, 1], [2, 3], range(4, 9))

# ---- Движок: плюсы и минусы ----
s = slide_pros_cons("Подход 3: движок IDE",
    ["Полнота: находит все ссылки, включая другие языки и конфигурацию",
     "Конфликты до правки: проверяет, куда приведёт вызов после правки, и показывает список конфликтов до применения",
     "Атомарность: поиск, проверка и правка – одна операция; один откат на всю"],
    ["Нужна запущенная IDE с индексами и кэшами",
     "Индексация занимает время; в этот период рефакторинг недоступен",
     "Применение правки – под блокировкой на запись",
     "Worktree, CI, удалённые машины – стандартного способа нет",
     "Runtime-рефлексия, связи через SQL, шаблоны, HTTP – вне модели кода"])
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
PX, PY, PW = 0.49, 1.2, 6.2                      # картинка на слайде
SC = PW / 2570                                   # дюймов на пиксель кропа (crop origin 30,260 в screenshot1.png)
def px(x, y):
    return PX + (x - 30) * SC, PY + (y - 260) * SC
picture(s, "img/psi-viewer.png", PX, PY, w=PW)   # 2570×1520 → 6.2×3.67
def outline(x1, y1, x2, y2, col=GREEN):
    ax, ay = px(x1, y1); bx, by = px(x2, y2)
    r = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(ax - 0.04), Inches(ay - 0.04), Inches(bx - ax + 0.08), Inches(by - ay + 0.08))
    r.adjustments[0] = 0.3; r.fill.background(); r.line.color.rgb = col; r.line.width = Pt(2); r.shadow.inherit = False
    return r
outline(815, 492, 930, 557)                      # токен lint в коде
outline(585, 1398, 1105, 1462)                   # узел REFERENCE_EXPRESSION
nx, ny = px(1105, 1430)
card(s, 6.75, 3.35, 2.75, 1.0, line=GREEN, size=10.5, font=MONO, gap=1, margin=0.12, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("RuleExtensions.kt", {"size": 9.5, "color": GREY, "font": FONT}),
    ("fun Rule.lint(content, …)", {"color": GREEN, "bold": True}),
])
dashed_arrow(s, nx + 0.06, ny, 6.7, 3.85)
textbox(s, 4.3, 3.35, 2.3, 0.4, "ссылка на объявление", size=10, color=GREEN, align=PP_ALIGN.CENTER)
anim(s, [0], [1, 2], [3, 4, 5])
notes(s, "Скриншот PSI Viewer IntelliJ на BracesOnIfStatementsSpec.kt:2213. Дерево – синтаксис, его видит и парсер. Ссылка – то, чего в тексте нет: узел lint знает, к какому объявлению относится; "
         "выбор зависит от типа receiver и импортов. В настоящем дереве у этой строки 15 узлов; индексы хранят, где какое имя встречается, – обход ссылок быстрый.")

# ---- Полнота ----
s = slide_blank("Подход 3: движок IDE", "Полнота")
_i = lambda: len(s.shapes) - 3
def codebox(x, y, w, text, col=WHITE, size=11.5):
    return card(s, x, y, w, 0.5, line=col, size=size, font=MONO, margin=0.12, anchor=MSO_ANCHOR.MIDDLE,
                text=[(text, {"color": col if col is not WHITE else WHITE, "bold": col is GREEN})])
g_k = _i()
card(s, 0.49, 1.25, 5.9, 3.9, line=LILAC, width=1.0, size=11, margin=0.15, text=[("Kotlin", {"color": LILAC, "bold": True})])
codebox(0.75, 1.75, 3.6, "fun Rule.lint(content: String)", GREEN)
for y, t in [(2.65, "subject.lint(code)"), (3.35, "lint(\"fun f() { $code }\")"), (4.05, "import …test.lint")]:
    codebox(3.0, y, 3.1, t)
    dashed_arrow(s, 2.95, y + 0.25, 2.2, 2.3)
g_j = _i()
card(s, 6.6, 1.25, 2.9, 1.75, line=LILAC, width=1.0, size=11, margin=0.15, text=[("Java", {"color": LILAC, "bold": True})])
codebox(6.75, 1.75, 2.6, "RuleExtensionsKt.lint(r, c)", size=10)
dashed_arrow(s, 6.7, 2.0, 4.4, 2.0)
g_x = _i()
card(s, 6.6, 3.2, 2.9, 1.95, line=LILAC, width=1.0, size=11, margin=0.15, text=[("XML", {"color": LILAC, "bold": True})])
codebox(6.75, 3.7, 2.6, "<bean class=\"…Rule\"/>")
dashed_arrow(s, 6.7, 3.95, 4.1, 2.3)
anim(s, range(g_k, g_j), range(g_j, g_x), list(range(g_x, _i())))
notes(s, "Объявление – одно, ссылки на него – из разных файлов и разных языков: модули Kotlin, Java, XML дают свои узлы и свои ссылки, движок обходит их одинаково. "
         "Пример Java и XML – иллюстрация возможности: в detekt функцию из Java не вызывают. Ссылка – слой поверх дерева: плагин может добавить связь туда, где компилятор её не видит. "
         "Индексы – почему обход быстрый. На некомпилирующемся коде – частично: неразрешённые ссылки помечаются отдельно.")

# ---- Корректность ----
s = slide_blank("Подход 3: движок IDE", "Корректность")
card(s, 0.72, 1.39, 8.56, 0.6, line=PURPLE, text="Правка примеряется на копии в памяти: конфликты видны до того, как тронут файл",
     size=13.5, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
steps = [("Копия вызова", "в памяти, файл не меняется"), ("Новое имя", "на копии"),
         ("Разрешение", "заново: какой метод?"), ("Другой метод", "→ конфликт")]
for i, (t, sub) in enumerate(steps):
    card(s, X4N[i], 2.29, 1.67, 0.8, line=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
         text=[(t, {"size": 12, "bold": True}), (sub, {"size": 10.5, "color": GREY})])
    if i < 3:
        arrow(s, X4N[i] + 1.8, 2.62)
card(s, 0.72, 3.39, 3.7, 1.7, line=WHITE, size=11.5, gap=6, text=[
    ("ДО RENAME", {"size": 10.5, "color": GREY, "bold": True}),
    ("rule.compileAndLint(code)", {"font": MONO, "size": 12}),
    ("rule.lint(code)  – другая функция, без компиляции", {"font": MONO, "size": 12, "color": GREEN}),
])
card(s, 5.58, 3.39, 3.7, 1.7, line=WHITE, size=11.5, gap=6, text=[
    ("ПОСЛЕ RENAME", {"size": 10.5, "color": GREY, "bold": True}),
    ("rule.lint(code)", {"font": MONO, "size": 12}),
    ("rule.lint(code)  – неразличимы; 78 вызовов меняют смысл", {"font": MONO, "size": 12, "color": RED}),
])
arrow(s, 4.7, 4.14, w=0.6, h=0.16)
anim(s, [0], range(1, 8), [8, 9, 10])
notes(s, "Механизм для метода: копия выражения-вызова, новое имя, повторный резолв. На кейсе – конфликт объявления.")

# ---- Таблица ----
s = slide_blank("Сравнение", "Результат и цена")
compare_table(s)
notes(s, "Не читать вслух. У текста нет символа; у LSP символ есть, но только rename, один язык и конфликты как повезёт; у движка – модель проекта и проверка до правки. "
         "Ниша текста – некомпилирующийся код и нулевая цена входа. Строка контекста – ответ на «сколько стоят описания инструментов».")

# ---- Bias агентов (a): Codex и Claude Code ----
s = slide_blank(CH_APPR, "Bias агентов")
source_line(s, 0.49, 1.15, 9.0, "Codex CLI – системный промпт")
picture(s, "img/codex-prompt.png", 0.49, 1.45, w=7.6)                # 1800×464 → 7.6×1.96
source_line(s, 0.49, 3.55, 9.0, "Claude Code – системный промпт")
picture(s, "img/claude-code-bypass-issue.png", 0.49, 3.85, w=7.6)    # 1640×329 → 7.6×1.52
notes(s, "Это инструкция агента, а не выбор модели: та же модель в другом агенте инструменты вызывает. Единственный инструмент правки у терминальных агентов – замена строки или патч. "
         "Источники: github.com/openai/codex, codex-rs/core/gpt_5_codex_prompt.md (тот же текст в промптах для gpt-5.2); блок bypass/auto – по issue anthropics/claude-code #88475 (версии 2.1.228+). "
         "Документация Claude Code: на macOS и Linux Glob и Grep убраны из набора по умолчанию – поиск через find и grep в Bash. OpenCode: описание инструмента Edit предлагает replaceAll, «if you want to rename a variable».")

# ---- Bias агентов (b): Cursor ----
s = slide_blank(CH_APPR, "Bias агентов")
source_line(s, 0.49, 1.15, 9.0, "Форум Cursor, июль 2026 – ответы сотрудников Cursor")
picture(s, "img/cursor-forum-colin.png", 0.49, 1.45, w=9.0)          # 2340×534 → 9.0×2.05
picture(s, "img/cursor-forum-maximep.png", 0.49, 3.65, w=9.0)        # 2340×382 → 9.0×1.47
notes(s, "Не промпт, а продуктовое решение: Cursor убрал семантический индекс, потому что модели справляются grep'ом. forum.cursor.com/t/…/165899")

# ---- Bias агентов (c): Serena ----
s = slide_blank(CH_APPR, "Bias агентов")
textbox(s, 0.49, 1.15, 9.0, 0.25, "Документация Serena – MCP-сервера с инструментами LSP для агентов, пост опубликован в октября 2026", size=11, color=GREY)
picture(s, "img/serena-docs-attention.png", 0.49, 1.52, h=3.75)      # 1640×1082 → 5.68×3.75
card(s, 6.4, 1.52, 3.1, 3.75, line=WHITE, size=12, gap=10, anchor=MSO_ANCHOR.MIDDLE, text=[
    "«Описания встроенных инструментов Claude Code – почти 16k токенов; пользователь их не меняет»",
    "«Почти невозможно убедить Opus 4.7 пользоваться Serena»",
    ("Авторы дают отдельный системный промпт против этого смещения", {"color": GREY}),
])
notes(s, "Переход к LSP: Serena даёт агенту LSP, и её авторы пишут, что модель к ней сама не идёт. Демо LSP – на Claude Code + Serena. oraios.github.io/serena/02-usage/030_clients.html")

# ---- Бенчмарк по моделям ----
s = slide_blank(CH_APPR, "Bias моделей")
card(s, 0.49, 1.4, 4.4, 3.7, line=AMBER, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER,
     text=[("TODO: бенчмарк по моделям", {"size": 16, "color": AMBER, "bold": True})])
card(s, 5.1, 1.4, 4.4, 3.7, line=WHITE, size=12, gap=8, text=[
    ("Пилот AgentConnect, август 2026", {"size": 13, "bold": True, "color": LILAC, "space_after": 8}),
    ("Модели Claude, LSP доступен как инструмент:", {"color": GREY}),
    ("•", "выбирали LSP в 0–6 % задач на поиск места в коде"),
    ("•", "в 3 % задач на переименование"),
    ("•", "принудительный LSP снизил долю успешных задач со 100 % до 89 %"),
    ("agentconnect.md/blog/grep-beat-lsp-harness", {"size": 10.5, "color": GREY}),
])
anim(s, [0], [1])
notes(s, "Оговорки – только здесь: пилот, 2–3 прогона на ячейку, rename через LSP не тестировали. Тезис статьи – «agent capability = model × harness»: "
         "модель сама к LSP не идёт, а если заставить без доработки агента – становится хуже.")

slide_plan(done={CH_INTRO, CH_APPR}, current="MCP-серверы для рефакторинга")

# ---- MCP ----
s = slide_blank("Готовые решения", "MCP-серверы для рефакторинга")
textbox(s, 0.49, 1.38, 9.0, 0.26,
        "MCP (Model Context Protocol) – протокол общения LLM с инструментами. "
        "MCP-сервер – сервер, отвечающий на запросы LLM по MCP",
        size=12, color=GREY)
card(s, 0.49, 2.0, 4.4, 3.1, line=WHITE, size=11.5, gap=4, text=[
    ("ВСТРОЕННЫЙ MCP-СЕРВЕР INTELLIJ IDEA", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 8}),
    ("•", "Встроен с версии 2025.2; сервер включается в настройках"),
    ("•", "Из рефакторингов – один инструмент: rename_refactoring"),
    ("•", "Параметры: путь в проекте, имя символа, новое имя – на перегрузках неоднозначно"),
    ("•", "Move и Safe Delete отсутствуют"),
])
card(s, 5.1, 2.0, 4.4, 3.1, line=WHITE, size=11.5, gap=4, text=[
    ("VEAI MCP", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 8}),
    ("•", "Инструменты рефакторинга IntelliJ, доступные агенту по MCP"),
    ("•", "Rename – переименование символа"),
    ("•", "Move – перенос символа в другой файл или пакет"),
    ("•", "Safe Delete – удаление с проверкой использований"),
])
anim(s, [0], [1], [2])

slide_plan(done={CH_INTRO, CH_APPR, "MCP-серверы для рефакторинга"}, current="Итоги")

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

# ---------- анимация: только там, где по клику появляется аргумент ----------
for _sl in prs.slides:
    _spec = ANIM.get(_sl.slide_id)
    if _spec:
        _content = list(_sl.shapes)[3:]
        animate(_sl, [[_content[i] for i in g] for g in _spec])

# ---------- язык текста: без lang="ru" PowerPoint проверяет орфографию как английскую ----------
from pptx.oxml.ns import qn as _qn
for _sl in prs.slides:
    for _r in _sl._element.iter(_qn("a:rPr")):
        _r.set("lang", "ru-RU")
    for _e in _sl._element.iter(_qn("a:endParaRPr")):
        _e.set("lang", "ru-RU")

prs.save(OUT)
print("слайдов:", len(prs.slides), "→", OUT)
