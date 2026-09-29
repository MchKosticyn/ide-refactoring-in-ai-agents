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


def card(slide, x, y, w, h, line=WHITE, width=2.25, fill=None, text=None, size=12,
         color=WHITE, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, gap=None,
         margin=0.2, font=FONT):
    """Скруглённая плашка в стиле шаблона. Текст – внутри самой фигуры.
    text – строка или список абзацев (см. _fill_tf)."""
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                               Inches(w), Inches(h))
    s.adjustments[0] = 0.08
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
    s.shapes.add_picture("veai_logo.png", Inches(0.4), Inches(0.4), width=Inches(1.01))
    textbox(s, 0.4, 2.95, 9.0, 0.4, speaker, size=14, color=GREY)
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


def slide_demo(kicker, title, steps, metrics=None, sub=None, actor="ДЕЙСТВИЯ АГЕНТА"):
    """Демо: слева область под видео, справа – метрики и шаги."""
    s = slide_blank(kicker, title)
    paras = [("TODO: видео", {"size": 12, "color": AMBER, "bold": True, "space_after": 8})]
    if sub:
        paras.append((sub, {"size": 12, "color": GREY}))
    card(s, 0.49, 1.35, 5.9, 3.8, line=AMBER, text=paras)
    y = 1.35
    if metrics:
        mh = 0.42 + 0.27 * len(metrics)
        card(s, 6.6, y, 2.9, mh, line=PURPLE, size=11, gap=2, margin=0.15,
             text=[("МЕТРИКИ", {"size": 10, "color": LILAC, "bold": True, "space_after": 4})]
                  + [(f"{k}: {v}", {"color": AMBER if "TODO" in v else WHITE}) for k, v in metrics])
        y += mh + 0.12
    card(s, 6.6, y, 2.9, 5.15 - y, line=WHITE, size=11, gap=5, margin=0.15,
         text=[(actor, {"size": 10, "color": LILAC, "bold": True, "space_after": 6})]
              + [(f"{i + 1}.  {st}", {}) for i, st in enumerate(steps)])
    return s


RESULT_ASPECTS = ["Найдены все места", "Правки только по делу", "Поведение сохранено"]


def result_table(s, statuses, banner=None):
    """Разбор демо: три строки. statuses – список (ok: bool, пояснение). Красный только если плохо."""
    y = 1.45
    for name, (ok, text) in zip(RESULT_ASPECTS, statuses):
        col = AMBER if ok is None else (GREEN if ok else RED)
        sym = "?" if ok is None else ("✓" if ok else "✗")
        # заголовок аспекта – отдельный текст слева (без рамки), статус и пояснение – внутри плашки
        textbox(s, 0.49, y, 2.4, 0.95, name, size=14, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        card(s, 2.95, y, 6.55, 0.95, line=col, anchor=MSO_ANCHOR.MIDDLE,
             text=[(f"{sym}  {text}", {"size": 12})], color=WHITE)
        # цвет значка: первый символ – отдельным run
        p = s.shapes[-1].text_frame.paragraphs[0]
        r0 = p.runs[0]
        full = r0.text
        r0.text = sym
        r0.font.size = Pt(18); r0.font.bold = True; r0.font.color.rgb = col
        r1 = p.add_run(); r1.text = full[len(sym):]
        r1.font.name = FONT; r1.font.size = Pt(12); r1.font.color.rgb = WHITE
        y += 1.08
    if banner:
        bottom_banner(s, banner, y=4.75, size=13)


def slide_pros_cons(kicker, pros, cons, note=None):
    s = slide_blank(kicker, "Плюсы и минусы")
    card(s, 0.49, 1.5, 4.4, 3.3, line=GREEN, size=13.5, gap=9,
         text=[("Плюсы", {"size": 17, "bold": True, "space_after": 12})] + [("•", p) for p in pros])
    card(s, 5.1, 1.5, 4.4, 3.3, line=RED, size=13.5, gap=9,
         text=[("Минусы", {"size": 17, "bold": True, "space_after": 12})] + [("•", c) for c in cons])
    if note:
        notes(s, note)
    return s


PLAN = [
    ("Введение", ["Что такое рефакторинг", "Зачем это QA", "Кейс", "Человек в IDE", "Термины"]),
    ("Подходы к рефакторингу в агентах", ["Подход 1: текстовый", "Подход 2: LSP", "Подход 3: движок IDE", "Сравнение"]),
    ("MCP для рефакторинга", None),
    ("Советы пользователям агентов", None),
]


def slide_plan(done=(), current=None, expand=None):
    """План доклада с прогрессом – один текстовый блок.
    done – пройденные пункты (зелёным), current – текущий (жирным), expand – глава, чьи подпункты показать."""
    s = slide_blank(TALK_SHORT, "План доклада")
    paras = []
    for title, subs in PLAN:
        col = GREEN if title in done else WHITE
        paras.append(("•", title, {"size": 17, "color": col, "bold": title == current, "space_after": 8}))
        if subs and title == expand:
            for sub in subs:
                c = GREEN if sub in done else WHITE
                paras.append(("–", sub, {"size": 14, "color": c, "bold": sub == current, "space_after": 6}))
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.4), Inches(8.8), Inches(3.8))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for mark, text, st in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(st["space_after"])
        _set_bullet(p)
        pPr = p._p.get_or_add_pPr()
        if mark == "–":
            pPr.set("marL", str(Emu(Inches(0.7))))
        r = p.add_run()
        r.text = text
        r.font.name = FONT
        r.font.size = Pt(st["size"])
        r.font.color.rgb = st["color"]
        r.font.bold = st["bold"]
    return s


def three_cards(s, titles, bodies, y=1.45, h=2.7, numbered=True, line=WHITE):
    xs = [0.49, 3.57, 6.65]
    for i, (t, b) in enumerate(zip(titles, bodies)):
        paras = []
        if numbered:
            paras.append((str(i + 1), {"size": 34, "color": PURPLE, "bold": True, "space_after": 2}))
        paras.append((t, {"size": 15, "bold": True, "space_after": 8}))
        paras.append((b, {"size": 11.5}))
        card(s, xs[i], y, 2.85, h, line=line if not isinstance(line, list) else line[i],
             text=paras, margin=0.25)


def bottom_banner(s, text, y=4.4, size=15):
    card(s, 1.2, y, 7.6, 0.68, line=PURPLE, text=text, size=size, bold=True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)


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


# =====================================================================
# СЛАЙДЫ  (по «План подробный.md», редакция 2026-09-29: кейс detekt)
# =====================================================================

CH_INTRO = "Введение"
CH_APPR = "Подходы к рефакторингу в агентах"
MONO = "Menlo"

# ---- 1. Титул ----
s = slide_title(["Использование инструментов", "рефакторинга IDE в ИИ-агентах"],
                "Спикер: Михаил Костицын")

# ---- 2. О себе ----
s = slide_blank(TALK_SHORT, "О себе")
card(s, 0.49, 1.4, 9.0, 2.2, line=PURPLE, size=14, gap=8, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("•", "Ведущий разработчик Veai – ИИ-агента для разработчиков"),
    ("•", "Семь лет в разработке, значительная часть – статический анализ программ и формальная верификация"),
    ("•", "С 2023 года – разработка ИИ-агентов"),
])
card(s, 0.49, 3.85, 9.0, 0.9, line=WHITE, size=13, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("Эксперт доклада: Максим – TODO фамилия, роль", {"color": AMBER}),
    ("Вместе делали инструменты рефакторинга в Veai и демо для этого доклада", {"size": 12, "color": GREY}),
])
notes(s, "Оргкомитет: эксперта упомянуть допустимо и уместно.")

# ---- план: старт ----
slide_plan(current="Что такое рефакторинг", expand=CH_INTRO)

# ---- 3. Рамка ----
s = slide_blank(CH_INTRO, "Рефакторинг – преобразование с гарантиями")
card(s, 0.72, 1.3, 8.56, 0.62, line=PURPLE, text="Рефакторинг – изменение структуры кода без изменения поведения программы",
     size=15, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
card(s, 0.72, 2.1, 4.1, 1.45, line=WHITE, size=12, gap=6, text=[
    ("Подзадача 1: найти все места", {"size": 15, "bold": True, "space_after": 8}),
    "Нужна полнота. 97 % не годится",
    ("78 вызовов × 97 % = 2 пропуска", {"color": LILAC, "bold": True}),
])
card(s, 5.18, 2.1, 4.1, 1.45, line=WHITE, size=12, gap=6, text=[
    ("Подзадача 2: применить", {"size": 15, "bold": True, "space_after": 8}),
    "Нужна корректность: правка не меняет, к какому методу привязан вызов, и не задевает чужое",
])
card(s, 0.72, 3.75, 8.56, 1.4, line=LILAC, size=12, gap=4, text=[
    ("ПО ЧЕМУ СУДИМ КАЖДЫЙ ПРОХОД", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 6}),
    ("•", "Результат: найдены все места · правки только по делу · поведение сохранено"),
    ("•", "Цена: время · стоимость · ходы · токены · стартовый контекст"),
])
notes(s, "Каркас для всех разборов. Одни и те же аспекты и метрики для человека и агентов. Не обсуждаем: RAG; «зачем агент, если Shift+F6» – Q&A.")

# ---- 4. Зачем это QA ----
s = slide_blank(CH_INTRO, "Зачем это QA")
three_cards(s, ["Миграции", "Структура", "Тестовый API"], [
    "Selenium → Playwright, JUnit 4 → 5, TestNG → JUnit. Тысячи вызовов, десятки файлов",
    "Spaghetti → PageObject; одна большая фикстура → композиция; аудит и переписывание по его итогам",
    "Хелперы, DSL, базовые классы тестов. Меняется контракт, которым написаны все тесты",
], y=1.35, h=2.2, numbered=False)
card(s, 0.49, 3.75, 9.0, 0.85, line=RED, size=12, gap=3, text=[
    ("ПОЧЕМУ БОЛЬНО", {"size": 10.5, "color": RED, "bold": True, "space_after": 4}),
    "Сотни мест; часть связей не выглядит как вызов; зелёные тесты – не доказательство, что поведение то же: меняется сам тестовый код",
])
bottom_banner(s, "Хочется отдать агенту. Первый опыт обычно разочаровывает – разберём, почему", y=4.75, size=13)
notes(s, "Примеры поводов – с митинга оргкомитета. Ручной способ есть (IDE), но долгий. Переход: берём небольшой реальный кейс, чтобы посчитать, а не рассуждать.")

# ---- 5. Кейс: проект и задача ----
s = slide_blank(CH_INTRO, "Кейс: detekt, PR #7873")
textbox(s, 0.49, 1.25, 9.0, 0.5,
        "detekt – статический анализатор Kotlin. detekt-test – тестовый DSL, которым авторы правил пишут спеки",
        size=12.5, color=GREY)
card(s, 0.49, 1.8, 9.0, 1.45, line=WHITE, size=11.5, font=MONO, gap=2, text=[
    ("RuleExtensions.kt", {"size": 10, "color": GREY, "font": FONT, "space_after": 6}),
    ("fun Rule.compileAndLint(content, compilerResources)", {"color": LILAC}),
    ("    // под флагом CI компилирует сниппет, потом линтит", {"color": GREY}),
    ("fun Rule.lint(content, compilerResources)", {"color": LILAC}),
    ("    // только линтит: для сниппетов, которые не компилируются", {"color": GREY}),
])
card(s, 0.49, 3.4, 9.0, 1.75, line=PURPLE, size=12, gap=4, text=[
    ("ЗАДАЧА АВТОРА PR (2025-02)", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 6}),
    ("«rename compileAndLint to lint so the function is always the same AND if you … doesn't want to compile "
     "the snippet … you can set .lint(code, compile = false)»", {"italic": True}),
    ("Реальный открытый проект, реальный PR. Можно повторить", {"size": 11, "color": GREY}),
])
notes(s, "Ловушки не называем – всплывут в разборах. Связка с «Зачем»: тестовый DSL вашего проекта устроен так же – пара хелперов, сотни вызовов, правило на дисциплине. См. demo-detekt.md.")

# ---- 6. Кейс: масштаб ----
s = slide_blank(CH_INTRO, "Кейс: масштаб")
nums = [("1428", "вызовов compileAndLint в 131 файле", WHITE),
        ("78", "вызовов старого lint в 17 файлах – им нужен compile = false", LILAC),
        ("36", "вызовов lint(ktFile) – перегрузка, не трогать", GREY),
        ("66", "вызовов FormattingRule.lint – другая функция с тем же именем, не трогать", GREY)]
xs4 = [0.49, 2.78, 5.07, 7.36]
for i, (n, t, c) in enumerate(nums):
    card(s, xs4[i], 1.4, 2.15, 2.6, line=c if c is not GREY else WHITE, margin=0.15, align=PP_ALIGN.CENTER,
         text=[(n, {"size": 40, "bold": True, "color": c, "space_after": 6}), (t, {"size": 11.5, "color": WHITE if c is not GREY else GREY})])
bottom_banner(s, "Задача – переименование. Всё интересное – в этих 78", y=4.3)
notes(s, "Только цифры. Числа посчитаны на pre-PR коммите a4ec32a2.")

# ---- 7. Человек в IDE: видео ----
s = slide_demo(CH_INTRO, "Человек в IDE",
               ["Change Signature на старом lint: параметр compile, в вызовах – false (78 мест)",
                "Rename compileAndLint → lint: конфликт «функция уже объявлена»",
                "Слить тела вручную: одна функция, if (compile && …)"],
               metrics=[("Время", "TODO"), ("Стоимость", "TODO (ставка × время)")],
               sub="IntelliJ IDEA. Инструменты IDE: Change Signature, Rename, Find Usages", actor="ДЕЙСТВИЯ")
notes(s, "Бейзлайн. Проверить при записи: как IDEA показывает конфликт Rename в занятое имя и что делает по «Continue»; "
         "подставляет ли Change Signature значение в вызов без точки (BracesOnIfStatementsSpec.kt:2212).")

# ---- 8. Человек в IDE: разбор ----
s = slide_blank(CH_INTRO, "Человек в IDE")
result_table(s, [
    (True, "Движок: все 78 вызовов, включая два без точки"),
    (True, "Только вызовы этого символа; FormattingRule.lint и lint(ktFile) не тронуты"),
    (True, "Конфликт показан до применения; compile = false подставлен явно"),
], banner="Время: TODO · Стоимость: TODO")
notes(s, "Rename в занятое имя – движок останавливает и показывает конфликт. Это то, что будем искать у агентов.")

# ---- 9. Термины ----
s = slide_blank(CH_INTRO, "Ключевые термины")
card(s, 0.49, 1.5, 2.9, 3.3, line=WHITE, size=11.5, gap=4, text=[
    ("КОНТЕКСТ", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 2}),
    ("что модель видит", {"size": 12, "color": GREY, "space_after": 12}),
    ("•", "системный промпт"), ("•", "история чата"), ("•", "прочитанные файлы"), ("•", "результаты вызовов инструментов"),
])
card(s, 3.75, 2.55, 2.5, 1.2, line=PURPLE, fill=RGBColor(0x1A, 0x12, 0x33), align=PP_ALIGN.CENTER,
     anchor=MSO_ANCHOR.MIDDLE, text=[("МОДЕЛЬ", {"size": 14, "bold": True}),
                                     ("генерирует сообщение\nили вызов инструмента", {"size": 10.5, "color": GREY})])
arrow(s, 3.42, 3.08, w=0.3, h=0.14)
arrow(s, 6.28, 3.08, w=0.3, h=0.14)
card(s, 6.61, 1.5, 2.9, 3.3, line=WHITE, size=11.5, gap=4, text=[
    ("ИНСТРУМЕНТЫ (TOOLS)", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 2}),
    ("что модель может вызвать", {"size": 12, "color": GREY, "space_after": 12}),
    ("•", "чтение файла"), ("•", "поиск по тексту"), ("•", "правка файла"), ("•", "запуск тестов"),
])
ret = s.shapes.add_shape(MSO_SHAPE.LEFT_ARROW, Inches(3.42), Inches(4.4), Inches(3.16), Inches(0.16))
ret.fill.solid(); ret.fill.fore_color.rgb = LILAC; ret.line.color.rgb = LILAC; ret.shadow.inherit = False
textbox(s, 3.42, 4.58, 3.16, 0.3, "результат вызова", size=10.5, color=GREY, align=PP_ALIGN.CENTER)
textbox(s, 0.49, 4.95, 9.0, 0.35,
        "ИИ-агент = модель + контекст + инструменты; работает циклом",
        size=11.5, color=GREY, align=PP_ALIGN.CENTER)
notes(s, "Только три термина. LSP, PSI, MCP – по ходу. Стартовый контекст как метрика опирается на это определение.")

# =============== план: введение пройдено =====
slide_plan(done={CH_INTRO}, current="Подход 1: текстовый", expand=CH_APPR)

# ---- 10. Демо 1: Claude Code ----
s = slide_demo(CH_APPR, "Подход 1: текстовый",
               ["grep по имени функции",
                "Чтение файлов",
                "sed / скрипт по файлам",
                "Точечные правки",
                "Отчёт"],
               metrics=[("Время", "13,7 мин"), ("Стоимость", "$4.41"), ("Ходы", "53"),
                        ("Токены", "64k out · 3,7M cache"), ("Стартовый контекст", "TODO")],
               sub="Claude Code, Opus. Тот же промпт – описание PR. Титры: агент, модель, дата")
notes(s, "Цифры прогона 2026-09-29 – перезаписать по факту записи. Промпт: описание PR, см. demo-detekt.md.")

# ---- 11. Разбор демо 1 ----
s = slide_blank(CH_APPR, "Подход 1: текстовый")
result_table(s, [
    (False, "2 из 78 вызовов пропущены: lint(\"fun f() { $code }\") внутри extension-функции – вызов без точки; агент искал «.lint(»"),
    (False, "Комментарий «not compileAndLint for performance reasons» стал «not lint for performance reasons»"),
    (False, "Два хелпера начали компилировать сотни сниппетов. Сборка зелёная – промах молчаливый"),
], banner="Проблема не в модели, а в инструментах")
notes(s, "Агент справился с основным: сам понял, что старые lint надо править до переименования, обошёл двойник. Промах – там, где имя не выглядит как вызов метода. "
         "Возврат к терминам: инструменты агента – поиск по тексту и правка текста.")

# ---- 12. Текстовый подход: что это ----
s = slide_blank(CH_APPR, "Подход 1: текстовый")
textbox(s, 0.49, 1.3, 9.0, 0.6,
        "Агент работает с кодом как с текстом: символов для него не существует. Так устроено большинство агентов",
        size=13, color=GREY)
labels = [("Задача", "текстом"), ("Поиск", "по тексту: grep, ripgrep"),
          ("Чтение", "найденных файлов"), ("Правка", "текстом"), ("Результат", "diff")]
xs5 = [0.49, 2.35, 4.21, 6.07, 7.93]
for i, (t, sub) in enumerate(labels):
    card(s, xs5[i], 2.15, 1.58, 0.95, line=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
         text=[(t, {"size": 13.5, "bold": True}), (sub, {"size": 10.5, "color": GREY})])
    if i < 4:
        arrow(s, xs5[i] + 1.62, 2.55, w=0.2, h=0.12)
card(s, 0.49, 3.4, 9.0, 1.7, line=PURPLE, size=12, gap=4, text=[
    ("МЕХАНИЗМЫ ПРАВКИ", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 8}),
    ("•", "Скрипт / regex по файлам – заменяет подстроку везде, где нашёл"),
    ("•", "Тул «правка фрагмента» – удалить / добавить / заменить; список мест – из поиска, каждое место – отдельный вызов"),
    ("•", "Перезапись файла целиком – модель генерирует новый файл"),
])

# ---- 13. Текстовый подход: плюсы и минусы ----
s = slide_pros_cons("Подход 1: текстовый",
    ["Не требует настройки",
     "Любой язык",
     "Работает на некомпилирующемся коде",
     "Есть в любом агенте"],
    ["Имя ≠ символ: одноимённая FormattingRule.lint, комментарий",
     "Символ ≠ имя: вызов без точки, неявный receiver",
     "Шум: частые идентификаторы дают сотни вхождений",
     "Правка не проверяет, что меняет; N мест – N операций"],
    note="Три плюса честно – последний вернётся в таблице как ниша.")

# ---- 14. Текстовый подход: где ломается ----
s = slide_blank("Подход 1: текстовый", "Где ломается")
card(s, 0.49, 1.35, 4.4, 2.6, line=RED, size=12, gap=6, text=[
    ("ПОЛНОТА", {"size": 10.5, "color": RED, "bold": True, "space_after": 6}),
    ("•", "Пропуски: символ без ожидаемой строки – вызов без точки, алиас, неявный receiver"),
    ("•", "Ложные попадания: строка без символа – одноимённый метод другого типа, комментарий"),
])
card(s, 5.1, 1.35, 4.4, 2.6, line=RED, size=12, gap=6, text=[
    ("КОРРЕКТНОСТЬ", {"size": 10.5, "color": RED, "bold": True, "space_after": 6}),
    ("•", "Правка не знает, что меняет: символ, строку или комментарий"),
    ("•", "N мест – N независимых операций; «прошло» всегда, даже если заменило не то"),
])
card(s, 0.49, 4.1, 9.0, 1.05, line=WHITE, anchor=MSO_ANCHOR.MIDDLE, size=11.5, font=MONO, text=[
    ("grep -rn '\\.lint('  → 180 строк.   lint(\"fun f() { $code }\")  – не найдено", {"color": LILAC}),
    ("Оба провала – из одной причины: подстрока не равна символу", {"font": FONT, "color": GREY}),
])
notes(s, "Вывод-переход: промах молчаливый, сборка зелёная, 53 хода и $4 – нужен инструмент, который видит символы, а не строки.")

# ---- 15. Почему агенты правят текстом ----
s = slide_blank(CH_APPR, "Почему агенты правят текстом")
quotes = [
    ("Codex CLI – системный промпт",
     "«prefer using rg … Do not use apply_patch … when scripting is more efficient (such as search and replacing a string across a codebase)»"),
    ("Claude Code – режим bypass / auto",
     "«make file changes with sed, heredocs, or short scripts, rather than using the dedicated Read, Edit, or Write tools»"),
    ("OpenCode – тул Edit",
     "«Use replaceAll … if you want to rename a variable»"),
]
y = 1.3
for t, q in quotes:
    card(s, 0.49, y, 9.0, 0.92, line=WHITE, size=11.5, margin=0.15, anchor=MSO_ANCHOR.MIDDLE,
         text=[(t, {"size": 10.5, "color": LILAC, "bold": True, "space_after": 2}), (q, {"italic": True})])
    y += 1.02
card(s, 0.49, 4.4, 9.0, 0.78, line=PURPLE, size=11, gap=1, margin=0.15, anchor=MSO_ANCHOR.MIDDLE, text=[
    "LSP / IDE-rename есть у Copilot CLI, JetBrains Junie, Serena",
    ("Смещение есть и у моделей: доля правильных вызовов тулов у Claude ниже примерно на 10 п.п. (бенчмарк Veai) – TODO согласовать", {"color": AMBER}),
])
notes(s, "Источники и permalink'и – research/agent-prompts-terminal.md. Не утверждать: что bypass-блок опубликован Anthropic (только issues #88475, #90599); "
         "что модели «переобучены». Это инструкция агента, а не выбор модели: та же модель в другом агенте тулы вызывает.")

# ---- 16. Veai ----
s = slide_blank(CH_APPR, "Veai")
card(s, 0.49, 1.4, 9.0, 1.6, line=PURPLE, size=13, gap=6, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("•", "ИИ-агент для разработчиков, работает внутри IDE JetBrains"),
    ("•", "Инструменты – те же, что у IDE: поиск по символам, рефакторинги, диагностика"),
    ("•", "Дальше – тот же кейс, тот же промпт, та же модель; меняем только набор инструментов"),
])
card(s, 0.49, 3.2, 9.0, 1.95, line=AMBER, size=12, gap=4, text=[
    ("TODO: скриншот панели тулов", {"size": 12, "color": AMBER, "bold": True, "space_after": 6}),
    ("•", "Конфигурация 2: текст + LSP, рефакторинг выключен"),
    ("•", "Конфигурация 3: текст + рефакторинг, LSP выключен"),
    ("•", "Стартовый контекст каждой – на слайдах демо"),
])
notes(s, "Полурекламный слайд: факты без эпитетов. Согласовать с компанией, что показывать из панели.")

# =============== план: подход 2 =====
slide_plan(done={CH_INTRO, "Подход 1: текстовый"}, current="Подход 2: LSP", expand=CH_APPR)

# ---- 17. Демо 2: LSP ----
s = slide_demo(CH_APPR, "Подход 2: LSP",
               ["Запрос references по символу",
                "Список мест от сервера",
                "Правки текстом по списку",
                "Rename по протоколу – отказ: «already declared»"],
               metrics=[("Время", "TODO"), ("Стоимость", "TODO"), ("Ходы", "TODO"),
                        ("Токены", "TODO"), ("Стартовый контекст", "TODO")],
               sub="Veai: LSP включён, рефакторинг выключен. Сервер: JetBrains Kotlin LSP. Титры: версия, дата")
notes(s, "Перед видео – одна фраза: LSP – сервер, который понимает язык; агент спрашивает его, где символ используется, и просит переименовать. Подробности после.")

# ---- 18. LSP: что это ----
s = slide_blank(CH_APPR, "Подход 2: LSP")
textbox(s, 0.49, 1.35, 9.0, 0.7,
        "LSP (Language Server Protocol) – протокол между редактором или агентом и сервером языка: отдельный процесс понимает язык, "
        "клиент спрашивает его по стандарту. Один сервер – один язык",
        size=13, color=GREY)
labels = [("Клиент", "редактор или агент"), ("Запрос", "«переименуй символ здесь»"),
          ("Сервер языка", "знает символы"), ("Правка", "список текстовых изменений; применяет клиент")]
xs = [0.72, 3.01, 5.33, 7.61]
for i, (t, sub) in enumerate(labels):
    card(s, xs[i], 2.3, 1.67, 1.0, line=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
         text=[(t, {"size": 14, "bold": True}), (sub, {"size": 10.5, "color": GREY})])
    if i < 3:
        arrow(s, xs[i] + 1.8, 2.73)
card(s, 0.72, 3.6, 8.56, 1.5, line=GREEN, size=12, gap=4, text=[
    ("Что даёт", {"size": 15, "bold": True, "space_after": 8}),
    ("•", "Переход к определению, поиск использований, переименование по символу"),
    ("•", "Один протокол для всех языков: rust-analyzer, gopls, pyright, Kotlin LSP"),
])
notes(s, "Первый подход, где есть понятие символа.")

# ---- 19. LSP: разбор ----
s = slide_blank(CH_APPR, "Подход 2: LSP")
result_table(s, [
    (True, "references на старом lint – 99 ссылок, включая оба вызова без точки"),
    (None, "По факту записи: правит ли агент по списку текстом"),
    (None, "По факту записи. Rename через протокол отказал: «Function 'lint' is already declared»"),
], banner="Заполнить по факту записи")
notes(s, "Rename отказал – лучше, чем молча применить. Текст ошибки модель прочитать может; проблема в том, что это выбор сервера, а не правило протокола, и продолжить после отказа нечем.")

# ---- 20. LSP: ограничения ----
s = slide_blank("Подход 2: LSP", "Ограничения протокола")
xs = [0.49, 5.1]
ys = [1.3, 2.95]
items = [
    ("Операции без параметров",
     "Один семантический рефакторинг – rename (позиция + имя). Остальное – code actions, которые предлагает сервер; "
     "передать «в какой модуль», «на какую сигнатуру» нечем"),
    ("Модель и файлы в разных процессах",
     "Сервер владеет моделью, клиент – файлами; о правках агента сервер узнаёт через уведомления. "
     "Сигнала «индекс актуален» нет; полноту references протокол не обещает"),
    ("Один сервер – один язык",
     "Ссылки из другого языка и из конфигурации недоступны; механизма композиции серверов в протоколе нет"),
    ("Конфликты – на усмотрение сервера",
     "Протокол не обязывает о них сообщать и не задаёт форму: ошибка строкой, молчаливое применение или пометка на правке. "
     "Параметра «применить несмотря на конфликт» нет"),
]
for i, (t, b) in enumerate(items):
    x = xs[i % 2]; y = ys[i // 2]
    card(s, x, y, 4.4, 1.5, line=RED, text=[(f"{i + 1}. {t}", {"size": 13, "bold": True, "space_after": 5}), (b, {"size": 11})])
bottom_banner(s, "Автор тула не может обещать модели то, что протокол не гарантирует", y=4.6, size=13)
notes(s, "Уровень протокола, источники – research/lsp-limitations.md. Оговорка: что сервер делает внутри ограничений – своё у каждого "
         "(clangd, gopls отказывают; rust-analyzer помечает; Kotlin LSP отказал нам). Наблюдение «2 → 79 ссылок» – иллюстрация. "
         "Вывод-переход: символы видит, оба вызова без точки нашёл; но операция одна – rename, конфликт как повезёт, один язык. Нужен движок.")

# =============== план: подход 3 =====
slide_plan(done={CH_INTRO, "Подход 1: текстовый", "Подход 2: LSP"}, current="Подход 3: движок IDE", expand=CH_APPR)

# ---- 21. Демо 3: движок ----
s = slide_demo(CH_APPR, "Подход 3: движок рефакторинга IDE",
               ["Тул Change Signature: параметр compile, false в вызовах",
                "Тул Rename: конфликт «уже объявлена»",
                "Решение по конфликту",
                "Применение; отчёт"],
               metrics=[("Время", "TODO"), ("Стоимость", "TODO"), ("Ходы", "TODO"),
                        ("Токены", "TODO"), ("Стартовый контекст", "TODO")],
               sub="Veai: рефакторинг включён, LSP выключен. Та же модель, тот же промпт. Титры: агент, модель, дата")

# ---- 22. Разбор демо 3 ----
s = slide_blank(CH_APPR, "Подход 3: движок рефакторинга IDE")
result_table(s, [
    (None, "Ожидание: все 78, включая вызовы без точки – по факту записи"),
    (None, "Ожидание: FormattingRule.lint и lint(ktFile) не тронуты – по факту записи"),
    (None, "Ожидание: конфликт до применения, compile = false подставлен – по факту записи"),
], banner="Та же модель, тот же промпт, другие тулы")
notes(s, "Переписать по факту записи. Метрики – рядом с человеком и Claude Code.")

# ---- 23. Движок: что это ----
s = slide_blank(CH_APPR, "Подход 3: движок рефакторинга IDE (IntelliJ)")
card(s, 0.49, 1.25, 9.0, 0.55, line=PURPLE, text="Модель решает «что», детерминированный инструмент – «как»",
     size=14, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
textbox(s, 0.49, 1.88, 9.0, 0.5,
        "Код представлен не текстом, а графом символов: символ – узел, ссылка – ребро к объявлению. «Все использования» – обход ссылок",
        size=12, color=GREY)
card(s, 0.49, 2.4, 4.4, 2.35, line=RED, size=11, font=MONO, gap=1, text=[
    ("КАК ВИДИТ GREP", {"size": 10.5, "color": RED, "bold": True, "font": FONT, "space_after": 6}),
    ("grep '\\.lint('  → 180 строк:", {"color": LILAC}),
    "  rule.lint(code)      Rule.lint",
    "  subject.lint(code)   FormattingRule.lint",
    "  rule.lint(ktFile)    перегрузка",
    ("  lint(\"fun f() …\")   – не видно", {"color": RED}),
])
card(s, 5.1, 2.4, 4.4, 2.35, line=GREEN, size=11, font=MONO, gap=1, text=[
    ("КАК ВИДИТ IDE", {"size": 10.5, "color": GREEN, "bold": True, "font": FONT, "space_after": 6}),
    ("Rule.lint(content)   99 ссылок", {"color": GREEN}),
    ("  в т. ч. 2 без точки", {"color": GREEN}),
    ("Rule.lint(ktFile)    другой символ", {"color": GREY}),
    ("FormattingRule.lint  другой символ", {"color": GREY}),
    ("Rule.compileAndLint  1558 ссылок", {"color": GREEN}),
])
textbox(s, 0.49, 4.82, 9.0, 0.4,
        "PSI (Program Structure Interface) – модель кода: дерево + ссылки между узлами. IDE держит её в памяти и синхронизирует с редактором",
        size=11.5, color=LILAC, align=PP_ALIGN.CENTER)
notes(s, "«Что / как»: агент решает, что переименовать и как поступить с конфликтом; движок делает работу и гарантирует результат. Рассматриваем только IntelliJ.")

# ---- 24. Движок: плюсы и минусы ----
s = slide_blank("Подход 3: движок IDE", "Плюсы и минусы")
card(s, 0.49, 1.35, 4.4, 3.75, line=GREEN, size=11.5, gap=4, text=[
    ("Плюсы", {"size": 17, "bold": True, "space_after": 12}),
    ("Полнота", {"size": 14, "bold": True, "space_after": 2}),
    ("Находит все ссылки, включая другие языки и конфигурацию", {"space_after": 12}),
    ("Конфликты до правки", {"size": 14, "bold": True, "space_after": 2}),
    ("Проверяет, что правка не изменит поведение; список конфликтов до применения", {"space_after": 12}),
    ("Атомарность", {"size": 14, "bold": True, "space_after": 2}),
    "Либо всё, либо ничего. Один откат на всю операцию",
])
card(s, 5.1, 1.35, 4.4, 3.75, line=RED, size=11.5, gap=5, text=[
    ("Минусы", {"size": 17, "bold": True, "space_after": 12}),
    ("•", "Нужна запущенная IDE с индексами и кэшами"),
    ("•", "Индексация занимает время; в этот период рефакторинг недоступен"),
    ("•", "Применение правки – под блокировкой на запись"),
    ("•", "Worktree, CI, удалённые машины – стандартного способа нет"),
    ("•", "Runtime-рефлексия, связи через SQL, шаблоны, HTTP – вне модели кода"),
    ("•", "Архитектурные изменения – не рефакторинг в смысле движка"),
])
notes(s, "Три гарантии – далее по слайду на каждую. Минусы – список фактов, без «но зато».")

# ---- 25. Полнота ----
s = slide_blank("Подход 3: движок IDE", "Полнота")
card(s, 0.72, 2.0, 2.2, 0.8, line=PURPLE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
     text=[("Rule.lint(content)", {"size": 12, "bold": True, "font": MONO}), ("объявление", {"size": 9.5, "color": GREY})])
refs = [("rule.lint(code)", 3.6, 1.5), ("lint(\"fun f() …\") – без точки", 3.6, 2.1), ("import …test.lint", 3.6, 2.7)]
for t, x, y in refs:
    card(s, x, y, 2.4, 0.5, line=WHITE, text=t, size=11, font=MONO, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05)
    arrow(s, 3.05, y + 0.18, w=0.5, h=0.14)
textbox(s, 0.72, 2.85, 2.2, 0.6, "ссылка: источник → цель;\nобратный поиск – все источники", size=10.5, color=GREY, align=PP_ALIGN.CENTER)
card(s, 6.4, 1.4, 3.1, 2.3, line=GREEN, size=11, gap=3, text=[
    ("ИСТОЧНИКИ ССЫЛОК", {"size": 10.5, "color": GREEN, "bold": True, "space_after": 6}),
    ("•", "Ссылка – слой поверх дерева"),
    ("•", "Плагины добавляют ссылки в другие языки и конфигурацию"),
    ("•", "Индексы: где какое имя встречается, кто кого наследует"),
    ("•", "Индексы и модель – в памяти IDE и кэшах на диске"),
])
textbox(s, 0.72, 3.5, 5.5, 0.5,
        "На некомпилирующемся коде – частично: неразрешённые ссылки помечаются отдельно",
        size=11, color=GREY)
card(s, 0.72, 4.05, 8.78, 1.05, line=RED, size=11.5, font=MONO, text=[
    ("ОГРАНИЧЕНИЕ", {"size": 10.5, "color": RED, "bold": True, "font": FONT, "space_after": 4}),
    "Class.forName(\"com.foo.Bar\")   – ссылка есть, переименуется",
    "Class.forName(prefix + name)   – runtime-рефлексия, не анализируется",
])
notes(s, "Ссылка – слой поверх дерева, поэтому плагин может добавить связь туда, где компилятор её не видит. Индексы – почему обход быстрый. Граница – сказать самому.")

# ---- 26. Корректность ----
s = slide_blank("Подход 3: движок IDE", "Корректность")
card(s, 0.72, 1.3, 8.56, 0.6, line=PURPLE, text="Проверки «имя занято» недостаточно – движок проверяет правку до применения",
     size=13.5, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
steps = [("Копия вызова", "оригинал не трогаем"), ("Новое имя", "на копии"),
         ("Разрешение", "заново: какой метод?"), ("Другой метод", "→ конфликт")]
xs = [0.72, 3.01, 5.33, 7.61]
for i, (t, sub) in enumerate(steps):
    card(s, xs[i], 2.1, 1.67, 0.8, line=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
         text=[(t, {"size": 12, "bold": True}), (sub, {"size": 10.5, "color": GREY})])
    if i < 3:
        arrow(s, xs[i] + 1.8, 2.43)
card(s, 0.72, 3.1, 3.7, 1.95, line=WHITE, size=11.5, gap=6, text=[
    ("ДО", {"size": 10.5, "color": GREY, "bold": True}),
    ("rule.compileAndLint(code)", {"font": MONO, "size": 12}),
    ("rule.lint(code)  – другая функция, без компиляции", {"font": MONO, "size": 12, "color": GREEN}),
])
card(s, 5.58, 3.1, 3.7, 1.95, line=WHITE, size=11.5, gap=6, text=[
    ("ПОСЛЕ RENAME", {"size": 10.5, "color": GREY, "bold": True}),
    ("rule.lint(code)", {"font": MONO, "size": 12}),
    ("rule.lint(code)  – неразличимы; 78 вызовов молча меняют смысл", {"font": MONO, "size": 12, "color": RED}),
])
arrow(s, 4.7, 3.98, w=0.6, h=0.16)
textbox(s, 0.72, 5.05, 8.56, 0.3, "Конфликт до применения: «функция lint уже объявлена». Текст применил и промолчал; LSP отказал строкой; движок показал, что именно конфликтует",
        size=11, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
notes(s, "Механизм для метода: копия выражения-вызова, новое имя, повторный резолв (RenameJavaMethodProcessor, advancedResolve). На кейсе – конфликт объявления.")

# ---- 27. Атомарность ----
s = slide_blank("Подход 3: движок IDE", "Атомарность")
card(s, 0.49, 1.4, 4.4, 2.3, line=WHITE, size=12, gap=8, text=[
    ("Одна команда", {"size": 15, "bold": True}),
    "Поиск использований, проверка конфликтов и правка – одна команда движка",
    ("Как транзакция в БД: либо все правки, либо ни одной", {"size": 11, "color": LILAC}),
])
card(s, 5.1, 1.4, 4.4, 2.3, line=WHITE, size=12, gap=8, text=[
    ("Один откат", {"size": 15, "bold": True}),
    "Глобальный откат на всю операцию, а не по файлу. Одно действие отмены",
    ("Не бывает «переименовано в 130 файлах из 133»", {"size": 11, "color": LILAC}),
])
card(s, 0.49, 3.95, 9.0, 1.15, line=PURPLE, size=12, text=[
    ("ПРОВЕРКА АКТУАЛЬНОСТИ", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 6}),
    "Код изменился между поиском и применением → рефакторинг не применяется; требуется повторный поиск",
])
notes(s, "Контраст с текстом: там N мест – N независимых операций; здесь одна. Строка IDE: «There were changes in code after usages have been found».")

slide_plan(done={CH_INTRO, "Подход 1: текстовый", "Подход 2: LSP", "Подход 3: движок IDE"}, current="Сравнение", expand=CH_APPR)

# ---- 28. Таблица ----
s = slide_blank("Сравнение", "Результат и цена")
cols = ["", "Человек в IDE", "Текст", "LSP", "Движок IDE"]
rows = [
    ("Найдены все места", ["✓", "✗ 2 из 78", "✓", "✓"]),
    ("Правки только по делу", ["✓", "✗ комментарий", "?", "✓"]),
    ("Поведение сохранено", ["✓", "✗ молча", "?", "✓"]),
    ("Время", ["TODO", "13,7 мин", "TODO", "TODO"]),
    ("Стоимость", ["TODO", "$4.41", "TODO", "TODO"]),
    ("Ходы", ["3 операции", "53", "TODO", "TODO"]),
    ("Стартовый контекст", ["–", "TODO", "TODO", "TODO"]),
]
x0, y0 = 0.49, 1.35
cw = [2.3, 1.7, 1.7, 1.6, 1.7]
rh = 0.47
x = x0
for i, c in enumerate(cols):
    if c:
        textbox(s, x, y0, cw[i], 0.4, c, size=12, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
    x += cw[i]
hline(s, x0, y0 + 0.42, sum(cw), color=PURPLE, width=1.2)
colmap = {"✓": GREEN, "✗": RED, "?": AMBER}
for r, (name, vals) in enumerate(rows):
    y = y0 + 0.5 + r * rh
    textbox(s, x0, y, cw[0], rh, name, size=11.5, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    x = x0 + cw[0]
    for i, v in enumerate(vals):
        col = colmap.get(v[0], AMBER if "TODO" in v else WHITE)
        textbox(s, x, y, cw[i + 1], rh, v, size=11.5, color=col, bold=v[0] in colmap, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x += cw[i + 1]
    if r == 2:
        hline(s, x0, y + rh, sum(cw), color=GREY, width=0.5)
notes(s, "Не читать вслух. Три подхода – в разных углах: у текста нет символа; у LSP символ есть, но только rename, один язык и конфликты как повезёт; "
         "у движка – модель проекта и проверка до правки. Ниша текста – некомпилирующийся код и нулевая цена входа. Строка контекста закрывает вопрос «сколько стоят описания тулов».")

slide_plan(done={CH_INTRO, CH_APPR}, current="MCP для рефакторинга")

# ---- 29. MCP ----
s = slide_blank("Готовые решения", "MCP для рефакторинга")
textbox(s, 0.49, 1.3, 9.0, 0.6,
        "MCP (Model Context Protocol) – протокол общения LLM с инструментами. "
        "MCP-сервер – сервер, отвечающий на запросы LLM по MCP",
        size=12, color=GREY)
card(s, 0.49, 2.0, 4.4, 3.1, line=WHITE, size=11.5, gap=4, text=[
    ("ВСТРОЕННЫЙ MCP-СЕРВЕР INTELLIJ IDEA", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 8}),
    ("•", "Встроен и включён по умолчанию с версии 2025.2"),
    ("•", "Из рефакторингов – один тул: rename_refactoring"),
    ("•", "Параметры: путь в проекте, имя символа, новое имя – на перегрузках неоднозначно"),
    ("•", "Extract Method, Change Signature, Inline – отсутствуют"),
])
card(s, 5.1, 2.0, 4.4, 3.1, line=WHITE, size=11.5, gap=4, text=[
    ("VEAI MCP", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 8}),
    ("•", "Тулы рефакторинга IntelliJ, доступные агенту по MCP"),
    ("TODO: список тулов", {"color": AMBER}),
])
notes(s, "Вывод следует из сравнения карточек. Перепроверить «включён по умолчанию с 2025.2».")

slide_plan(done={CH_INTRO, CH_APPR, "MCP для рефакторинга"}, current="Советы пользователям агентов")

# ---- 30. Советы ----
s = slide_blank("Советы", "Советы пользователям агентов")
card(s, 0.49, 1.7, 4.4, 2.0, line=GREEN, size=12.5, gap=10, text=[
    ("Рефакторинг – отдельным коммитом", {"size": 15, "bold": True}),
    "Отдельно от изменения поведения. Ревью читаемо, откат возможен",
])
card(s, 5.1, 1.7, 4.4, 2.0, line=GREEN, size=12.5, gap=10, text=[
    ("Регрессионные тесты как фиксация поведения", {"size": 15, "bold": True}),
    "Тесты фиксируют поведение до рефакторинга; прогон после подтверждает, что поведение не изменилось",
])

# ---- 31. Summary ----
s = slide_blank("Summary", "Главные мысли")
three_cards(s, ["Рефакторинг – преобразование с гарантиями", "Автоматическая проверяемость обязательна", "Инструмент рефакторинга уже есть – в IDE"], [
    "Полнота при поиске, корректность при применении",
    "Тесты, линтеры, компиляция – особенно при текстовых правках, где агент ошибается молча. В кейсе сборка была зелёной",
    "Полнота, конфликты до правки, атомарность. Модель решает «что», движок делает «как». Отдайте его агенту",
], y=1.45, h=3.0)
notes(s, "По фразе на тезис. Третий – призыв к действию.")

# ---- 32. Спасибо ----
s = prs.slides.add_slide(L_TEXT)
_clear_placeholders(s)
_chrome(s, TALK_SHORT)
textbox(s, 0.4, 1.6, 6.9, 2.0, ["Спасибо", "за внимание"], size=48, color=WHITE)
textbox(s, 0.4, 4.0, 6.0, 1.0, [("Вопросы?", {"size": 20, "color": LILAC, "bold": True})], size=20)
textbox(s, 5.17, 4.39, 4.42, 0.82, [
    ("Попробовать агента: veai.ru", {"bold": True}),
    ("Telegram-канал: @veai_devs", {"bold": True}),
], size=13, color=WHITE)

# ---- Backup: четыре принципа ----
s = slide_blank("Дополнительно", "Как отдать движок агенту: четыре принципа")
xs4 = [0.49, 2.78, 5.07, 7.36]
principles = [
    ("Адресация", "Не оффсеты и строки, а сигнатура и стабильный идентификатор символа. Отдельный тул поиска возвращает кандидатов с контекстом"),
    ("Гранулярность", "Набор тулов по видам рефакторинга с параметрами. Описание тула – промпт: когда применять, что ожидать"),
    ("preview → apply", "Сначала «что произойдёт» структурой: места, конфликты. Потом применение. Три состояния: применено / конфликт / устарело"),
    ("Ошибки как промпт", "Текст ошибки читает модель: что не так, где, из-за чего. «Refactoring failed» не даёт следующего шага; «символ неоднозначен, кандидаты: …» – даёт"),
]
for i, (t, b) in enumerate(principles):
    card(s, xs4[i], 1.4, 2.15, 3.7, line=WHITE, margin=0.15,
         text=[(t, {"size": 13.5, "bold": True, "space_after": 8}), (b, {"size": 11.5})])
notes(s, "Backup-слайд для Q&A.")


prs.save(OUT)
print("слайдов:", len(prs.slides), "→", OUT)
