"""Сборка презентации доклада «Не давайте ИИ-агенту рефакторить код руками»
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


def slide_demo(kicker, title, steps, sub=None):
    """Демо: слева область под видео, справа – шаги агента."""
    s = slide_blank(kicker, title)
    paras = [("TODO: видео", {"size": 12, "color": AMBER, "bold": True, "space_after": 8})]
    if sub:
        paras.append((sub, {"size": 12, "color": GREY}))
    card(s, 0.49, 1.35, 5.9, 3.8, line=AMBER, text=paras)
    card(s, 6.6, 1.35, 2.9, 3.8, line=WHITE, size=11.5, gap=8,
         text=[("ДЕЙСТВИЯ АГЕНТА", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 10})]
              + [(f"{i + 1}.  {st}", {}) for i, st in enumerate(steps)])
    return s


RESULT_ASPECTS = ["Найдены все места", "Правки только по делу", "Поведение сохранено"]


def result_table(s, statuses, banner=None):
    """Разбор демо: три строки. statuses – список (ok: bool, пояснение). Красный только если плохо."""
    y = 1.45
    for name, (ok, text) in zip(RESULT_ASPECTS, statuses):
        col = GREEN if ok else RED
        sym = "✓" if ok else "✗"
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
    ("Введение", None),
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
# СЛАЙДЫ  (по «План подробный.md»)
# =====================================================================

# ---- 1. Титул ----
s = slide_title(["Использование инструментов", "рефакторинга IDE в ИИ-агентах"],
                "Спикер: Михаил Костицын")

# ---- 2. О спикере ----
s = slide_blank(TALK_SHORT, "О себе")
card(s, 0.49, 1.4, 9.0, 2.2, line=PURPLE, size=14, gap=8, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("•", "Ведущий разработчик Veai – ИИ-агента для разработчиков"),
    ("•", "Семь лет в разработке, значительная часть – статический анализ программ и формальная верификация"),
    ("•", "С 2023 года – разработка ИИ-агентов"),
])

# ---- план: старт ----
slide_plan(current="Введение")

# ---- 3. Термины: схема ----
s = slide_blank("Введение", "Ключевые термины")
# контекст слева
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
# возврат результата: подпись внизу
# возврат результата: стрелка снизу от инструментов к контексту
ret = s.shapes.add_shape(MSO_SHAPE.LEFT_ARROW, Inches(3.42), Inches(4.4), Inches(3.16), Inches(0.16))
ret.fill.solid(); ret.fill.fore_color.rgb = LILAC; ret.line.color.rgb = LILAC; ret.shadow.inherit = False
textbox(s, 3.42, 4.58, 3.16, 0.3, "результат вызова", size=10.5, color=GREY, align=PP_ALIGN.CENTER)
textbox(s, 0.49, 4.95, 9.0, 0.35,
        "ИИ-агент = модель + контекст + инструменты; работает циклом",
        size=11.5, color=GREY, align=PP_ALIGN.CENTER)
notes(s, "Только три термина. LSP, PSI, MCP – по ходу. После демо – возврат: агент искал текстом и правил текстом, это были его инструменты.")

# ---- 6. Рамка ----
s = slide_blank("Идея", "Рефакторинг – преобразование с гарантиями")
card(s, 0.72, 1.35, 8.56, 0.7, line=PURPLE, text="Рефакторинг – изменение структуры кода без изменения поведения программы",
     size=15, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
card(s, 0.72, 2.3, 4.1, 1.75, line=WHITE, size=12, gap=6, text=[
    ("Подзадача 1: найти все места", {"size": 15, "bold": True, "space_after": 10}),
    "Нужна полнота. 99 % не годится",
    ("200 вызовов × 99 % = 2 пропуска", {"color": LILAC, "bold": True}),
])
card(s, 5.18, 2.3, 4.1, 1.75, line=WHITE, size=12, gap=6, text=[
    ("Подзадача 2: выполнить", {"size": 15, "bold": True, "space_after": 10}),
    "Нужна корректность. Правка не должна менять привязку вызовов, видимость, иерархию переопределений",
])
bottom_banner(s, "Модель решает «что», детерминированный инструмент – «как»", y=4.3)
notes(s, "Рамка – первое из трёх повторений (6, 18). Не обсуждаем: RAG, «зачем агент, если есть Shift+F6» – Q&A.")

# =============== план: введение пройдено =====
slide_plan(done={"Введение"}, current="Подход 1: текстовый", expand="Подходы к рефакторингу в агентах")

# ---- 4. Демо 1: текст ----
s = slide_demo("Подходы к рефакторингу в агентах", "Подход 1: текстовый",
               ["Получает задачу: переименовать метод",
                "Ищет вхождения имени по тексту",
                "Читает найденные файлы",
                "Правит каждое место текстом",
                "Отдаёт diff"],
               sub="Java/Kotlin-проект: перегрузка, метод с целевым именем, Java-переопределение, "
                   "вызов из Kotlin через свойство (user.name), строка лога, Spring XML")
notes(s, "См. demo-spec.md и «TODO внутреннее.md».")

# ---- 5. Разбор демо 1 ----
s = slide_blank("Подходы к рефакторингу в агентах", "Подход 1: текстовый")
result_table(s, [
    (False, "Пропущен вызов из Kotlin через синтаксис свойства (user.name) – подстроки getName нет. Сборка не проходит"),
    (False, "Изменены перегрузка getName(Locale) и строка лога – совпадение подстроки. Сборка проходит"),
    (False, "Вызов привязался к другому методу с тем же именем. Сборка и тесты проходят, поведение другое"),
], banner="Проблема не в модели, а в инструментах")
notes(s, "Возврат к терминам: агент искал текстом и правил текстом. Третья карточка – сборка и тесты проходят, поведение изменилось.")

# ---- Текстовый подход: что это ----
s = slide_blank("Подходы к рефакторингу в агентах", "Подход 1: текстовый")
textbox(s, 0.49, 1.3, 9.0, 0.6,
        "Агент работает с кодом как с текстом: символов для него не существует. Так устроено большинство агентов",
        size=13, color=GREY)
labels = [("Задача", "переименовать метод"), ("Поиск", "по тексту: grep, ripgrep"),
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
notes(s, "Что из себя представляет подход. Три механизма правки названы здесь; на слайде «Корректность» – только где ломаются.")

# ---- Текстовый подход: плюсы и минусы ----
s = slide_pros_cons("Подход 1: текстовый",
    ["Не требует настройки",
     "Любой язык",
     "Работает на некомпилирующемся коде",
     "Есть в любом агенте"],
    ["Полнота: имя ≠ символ – ложные попадания и пропуски",
     "Полнота: шум",
     "Корректность: правка не проверяет, что меняет",
     "Корректность: N мест – N независимых операций"],
    note="Top-down: здесь коротко, далее два слайда – полнота и корректность.")

# ---- Текстовый подход: полнота ----
s = slide_blank("Подход 1: текстовый", "Полнота")
card(s, 0.49, 1.35, 9.0, 2.7, line=RED, size=13, gap=8, text=[
    ("ГДЕ ЛОМАЕТСЯ", {"size": 10.5, "color": RED, "bold": True, "space_after": 8}),
    ("•", "Имя без символа: одноимённые вхождения в другом классе, в комментариях, в строках – ложные попадания"),
    ("•", "Символ без имени: алиасы импорта, обращение через свойство – пропуски"),
    ("•", "Шум: частые идентификаторы (run, config, init) дают сотни вхождений"),
])
card(s, 0.49, 4.2, 9.0, 0.95, line=WHITE, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("val label = user.name   // вызов getName(), подстроки \"getName\" нет", {"font": "Menlo", "size": 12}),
    ("Исследование: 71–75 % провалов текстового поиска в агенте – найденное место не выбрано из-за шума", {"size": 11.5, "color": GREY}),
])
notes(s, "Где ломается полнота: пропуски (символ без имени) и ложные попадания (имя без символа). Цифра – GrepRAG, arXiv 2601.23254, сверить.")

# ---- Текстовый подход: правка ----
s = slide_blank("Подход 1: текстовый", "Корректность")
card(s, 0.49, 1.35, 9.0, 2.7, line=RED, size=13, gap=8, text=[
    ("ГДЕ ЛОМАЕТСЯ", {"size": 10.5, "color": RED, "bold": True, "space_after": 8}),
    ("•", "Скрипт / regex: правит строки и комментарии; не проверяет, что заменённое – тот самый символ"),
    ("•", "Тул «правка фрагмента»: каждое место – отдельная операция без проверки; ошибка в списке мест – ошибка в правке"),
    ("•", "Перезапись файла: помимо целевой правки может изменить соседний код"),
])
card(s, 0.49, 4.2, 9.0, 0.95, line=WHITE, anchor=MSO_ANCHOR.MIDDLE, text=[
    ("Каждая правка – независимая текстовая операция без проверки результата", {"size": 12.5, "bold": True}),
    ("N мест – N независимых операций; ошибка не обнаруживается в момент применения", {"size": 11.5, "color": GREY}),
])
notes(s, "Где ломается корректность: правка не знает, что меняет; не проверяет привязку вызова; N мест – N независимых операций, состояние «7 из 9» возможно.")

# ---- LSP: демо ----
s = slide_demo("Подходы к рефакторингу в агентах", "Подход 2: LSP",
               ["F2 на объявлении метода",
                "Редактор отправляет запрос серверу",
                "Сервер (jdtls) ищет использования по символу",
                "Сервер возвращает список правок",
                "Редактор применяет правки"],
               sub="Тот же проект. VS Code + Java-сервер (jdtls). Без агента – переименование вызывается напрямую")

# ---- LSP: разбор ----
s = slide_blank("Подходы к рефакторингу в агентах", "Подход 2: LSP")
result_table(s, [
    (False, "Пропущены вызов из Kotlin через синтаксис свойства и Spring XML – сервер о них не знает"),
    (True, "Перегрузка getName(Locale) и строка лога не тронуты"),
    (False, "Вызов привязался к другому методу – сервер не сообщил о конфликте"),
], banner="Полнота в одном языке есть; за границей языка – нет; конфликты не сообщаются")
notes(s, "Карточки заполнены по ожиданию – переписать по факту записи.")
notes(s, "Контраст с демо 1: меньше лишнего, но пропуски там же, где нет символа для этого сервера.")

# ---- LSP: что это ----
s = slide_blank("Подходы к рефакторингу в агентах", "Подход 2: LSP")
textbox(s, 0.49, 1.35, 9.0, 0.7,
        "LSP – протокол между редактором и сервером языка: отдельный процесс понимает язык, "
        "редактор или агент спрашивает его по стандартному протоколу. Как правило, один сервер – один язык",
        size=13, color=GREY)
labels = [("Редактор", "или агент"), ("Запрос", "«переименуй символ здесь»"),
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
    ("•", "Один протокол для всех языков: rust-analyzer, gopls, jdtls, pyright"),
])
notes(s, "Первый подход, где есть понятие символа. Деталь схемы для атомарности: сервер возвращает текстовые правки, применяет клиент.")

# ---- LSP: плюсы и минусы ----
s = slide_pros_cons("Подход 2: LSP",
    ["Поиск по символу, а не по подстроке",
     "Один протокол для всех языков",
     "Есть в любом редакторе",
     "Не требует IDE"],
    ["Конфликты не передаются через протокол",
     "Параметров у рефакторинга нет",
     "Атомарность не гарантируется",
     "Межъязыковой рефакторинг невозможен"],
    note="Top-down: минусы – далее по карточке на каждый, затем пример.")

# ---- LSP: пределы ----
s = slide_blank("Подход 2: LSP", "Пределы")
xs = [0.49, 5.1]
ys = [1.3, 3.05]
items = [
    ("Протокол: конфликты не передаются",
     "Запрос на переименование – документ, позиция, новое имя. Ответ – готовая правка или ошибка строкой. "
     "Структурированного статуса конфликта у ответа нет"),
    ("Протокол: параметров нет",
     "Code action – команда без параметров. Спецификация приводит «Add or remove parameter» как пример, "
     "но не даёт способа передать, какой параметр и куда"),
    ("Клиент: атомарность не гарантируется",
     "Правка – список текстовых изменений, применяет клиент. Стратегия отката – декларация клиента. "
     "VS Code: транзакционно только для текста"),
    ("Клиент: один сервер на запрос",
     "Клиент берёт ответ первого сервера, остальные не опрашивает. Java-сервер не знает о Kotlin – "
     "межъязыковой рефакторинг невозможен"),
]
for i, (t, b) in enumerate(items):
    x = xs[i % 2]; y = ys[i // 2]
    card(s, x, y, 4.4, 1.6, line=RED, text=[(f"{i + 1}. {t}", {"size": 13, "bold": True, "space_after": 6}), (b, {"size": 11.5})])
textbox(s, 0.49, 4.75, 9.0, 0.45,
        "На некомпилирующемся коде сервер вправе отказать. Поведение зависит от реализации сервера",
        size=11.5, color=GREY)
notes(s, "Поиск у LSP – по символам, применение – текстовое, клиентом: по подзадаче 1 ближе к движку, по подзадаче 2 – к тексту. "
         "В заметках для Q&A: issue открыт с 2016, цитаты автора протокола, багрепорты 2024, «первый ответивший побеждает» в коде VS Code. "
         "См. notes-lsp-vs-psi.md, раздел 1.")

# ---- LSP: пример j → i ----
s = slide_blank("Подход 2: LSP", "Пример")
card(s, 0.49, 1.35, 5.2, 3.75, line=WHITE, size=12, font="Menlo", text=[
    ("JAVA, VS CODE: ПЕРЕИМЕНОВАНИЕ j → i", {"size": 10.5, "color": GREY, "bold": True, "font": FONT, "space_after": 10}),
    "void m() {", "  int i = 0;", "  class B {",
    ("    int j;              // переименуем в i", {"color": LILAC}),
    ("    void k() { j = i; }", {"color": RED}),
    ("    // после правки:  i = i", {"color": RED}),
    "  }", "}",
    ("Сборка и тесты проходят. Поведение изменилось", {"size": 12, "color": RED, "bold": True, "font": FONT, "space_after": 0}),
])
card(s, 5.9, 1.35, 3.6, 3.75, line=WHITE, size=13, gap=10, text=[
    ("ЧТО ПРОИЗОШЛО", {"size": 10.5, "color": GREY, "bold": True, "space_after": 12}),
    "Движок нашёл конфликт.",
    "Протокол его не передал – правка применена без предупреждения.",
    "В Eclipse IDE – диалог с предупреждением. Через LSP – без предупреждения",
])
notes(s, "Механизм (для Q&A): движок Eclipse вычислил конфликты с уровнями; LSP-обработчик прерывается только на «фатально»; "
         "функция сборки правки принимает только набор изменений – статус передать некуда. "
         "Устно: LSP не заменяет движок, а ограничивает доступ к нему.")

# ---- Движок: демо ----
s = slide_demo("Подходы к рефакторингу в агентах", "Подход 3: движок рефакторинга IDE",
               ["Получает ту же задачу",
                "Находит символ тулом поиска",
                "Вызывает тул переименования: preview",
                "Получает список мест и конфликт",
                "Применяет; сообщает о вхождении в строке"],
               sub="Тот же проект, тот же запрос, та же модель. Veai-агент")

# ---- Движок: разбор ----
s = slide_blank("Подходы к рефакторингу в агентах", "Подход 3: движок рефакторинга IDE")
result_table(s, [
    (True, "Все вызовы, включая Kotlin и Spring XML"),
    (True, "Строка лога не тронута; агент сообщил о вхождении в строке"),
    (True, "Конфликт «будет вызван другой метод» показан до применения"),
], banner="Та же модель, тот же запрос, другие тулы")
notes(s, "Карточки заполнены по ожиданию – переписать по факту записи.")
notes(s, "Не «ноль пропусков», а «ноль незамеченных пропусков»: про строку лога агент сообщил.")

# ---- Движок: что это (определение + PSI) ----
s = slide_blank("Подходы к рефакторингу в агентах", "Подход 3: движок рефакторинга IDE (IntelliJ)")
textbox(s, 0.49, 1.3, 9.0, 0.6,
        "Код представлен не текстом, а связанным графом символов: символ – узел, ссылка – ребро к объявлению. "
        "«Все использования» – обход ссылок, а не поиск подстроки",
        size=13, color=GREY)
card(s, 0.49, 2.0, 4.4, 2.75, line=RED, size=11, font="Menlo", text=[
    ("КАК ВИДИТ ТЕКСТОВЫЙ ПОИСК", {"size": 10.5, "color": RED, "bold": True, "font": FONT, "space_after": 8}),
    "class User {",
    ("  String getName() {...}", {"color": LILAC}),
    ("  String getName(Locale l) {...}", {"color": LILAC}),
    "}",
    ("user.getName();", {"color": LILAC}),
    ("log.info(\"getName called\");", {"color": LILAC}),
    ("order.getName();", {"color": LILAC}),
    ("val label = user.name  // Kotlin", {"color": GREY}),
])
card(s, 5.1, 2.0, 4.4, 2.75, line=GREEN, size=11, font="Menlo", text=[
    ("КАК ВИДИТ IDE", {"size": 10.5, "color": GREEN, "bold": True, "font": FONT, "space_after": 8}),
    ("User.getName()       объявление", {"color": GREEN}),
    ("User.getName(Locale) другой символ", {"color": GREY}),
    ("user.getName()  ──▶  ссылка", {"color": GREEN}),
    ("\"getName called\"     строка, не ссылка", {"color": GREY}),
    ("order.getName() ──▶  Order.getName()", {"color": GREY}),
    ("user.name       ──▶  ссылка (Kotlin)", {"color": GREEN}),
])
textbox(s, 0.49, 4.8, 9.0, 0.4,
        "PSI (Program Structure Interface) – модель кода: дерево + ссылки между узлами. IDE держит её в памяти и синхронизирует с редактором",
        size=11.5, color=LILAC, align=PP_ALIGN.CENTER)
notes(s, "Что из себя представляет подход. Вводим PSI. Рассматриваем только IntelliJ.")

# ---- Движок: плюсы и минусы ----
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
textbox(s, 0.49, 5.15, 9.0, 0.3, "TODO: worktree и CI – уточнить формулировки по практике", size=10.5, color=AMBER)
notes(s, "Три гарантии – далее по слайду на каждую. Минусы – список фактов, без «но зато».")

# ---- Полнота 2: ссылки ----
s = slide_blank("Подход 3: движок IDE", "Полнота")
card(s, 0.72, 2.0, 2.2, 0.8, line=PURPLE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05,
     text=[("getName()", {"size": 13, "bold": True}), ("объявление", {"size": 9.5, "color": GREY})])
refs = [("вызов в Java", 3.6, 1.5), ("вызов из Kotlin: user.name", 3.6, 2.1), ("имя свойства в Spring XML", 3.6, 2.7)]
for t, x, y in refs:
    card(s, x, y, 2.2, 0.5, line=WHITE, text=t, size=11, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.05)
    arrow(s, 3.05, y + 0.18, w=0.5, h=0.14)
textbox(s, 0.72, 2.85, 2.2, 0.6, "ссылка: источник → цель;\nобратный поиск – все источники", size=10.5, color=GREY, align=PP_ALIGN.CENTER)
card(s, 6.2, 1.4, 3.3, 2.3, line=GREEN, size=11, gap=3, text=[
    ("ИСТОЧНИКИ ССЫЛОК", {"size": 10.5, "color": GREEN, "bold": True, "space_after": 6}),
    ("•", "Ссылка – слой поверх дерева"),
    ("•", "Плагины добавляют ссылки в другие языки и конфигурацию"),
    ("•", "Индексы: где какое имя встречается, кто кого наследует"),
    ("•", "Индексы и модель – в памяти IDE и кэшах на диске"),
])
textbox(s, 0.72, 3.5, 5.3, 0.5,
        "На некомпилирующемся коде – частично: неразрешённые ссылки помечаются отдельно",
        size=11, color=GREY)
card(s, 0.72, 4.05, 8.78, 1.05, line=RED, size=11.5, font="Menlo", text=[
    ("ОГРАНИЧЕНИЕ", {"size": 10.5, "color": RED, "bold": True, "font": FONT, "space_after": 4}),
    "Class.forName(\"com.foo.Bar\")   – ссылка есть, переименуется",
    "Class.forName(prefix + name)   – runtime-рефлексия, не анализируется",
])
notes(s, "Ссылка – слой поверх дерева, поэтому плагин может добавить связь туда, где компилятор её не видит. Индексы – почему обход быстрый; цена – п. 11. Граница – сказать самому.")

# ---- Корректность ----
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
    ("account.getName(locale)", {"font": "Menlo", "size": 12}),
    ("привязан к getName(Object) – его переименуем", {"color": GREEN}),
])
card(s, 5.58, 3.1, 3.7, 1.95, line=WHITE, size=11.5, gap=6, text=[
    ("ПОСЛЕ", {"size": 10.5, "color": GREY, "bold": True}),
    ("account.getDisplayName(locale)", {"font": "Menlo", "size": 12}),
    ("привязался бы к существующему getDisplayName(Locale) – другой метод", {"color": RED}),
])
arrow(s, 4.7, 3.98, w=0.6, h=0.16)
textbox(s, 0.72, 5.05, 8.56, 0.3, "Конфликт: «После переименования будет вызван другой метод»",
        size=11.5, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
notes(s, "Ловушка 2 из demo-spec.md: текст применил молча, LSP применил молча, движок остановился. Существующий метод более специфичен, чем переименуемый – поэтому вызов пересаживается. Механизм – RenameJavaMethodProcessor, advancedResolve.")

# ---- Атомарность ----
s = slide_blank("Подход 3: движок IDE", "Атомарность")
card(s, 0.49, 1.4, 4.4, 2.3, line=WHITE, size=12, gap=8, text=[
    ("Одна команда", {"size": 15, "bold": True}),
    "Поиск использований, проверка конфликтов и правка – одна команда движка",
    ("Как транзакция в БД: либо все правки, либо ни одной", {"size": 11, "color": LILAC}),
])
card(s, 5.1, 1.4, 4.4, 2.3, line=WHITE, size=12, gap=8, text=[
    ("Один откат", {"size": 15, "bold": True}),
    "Глобальный откат на всю операцию, а не по файлу. Одно действие отмены",
    ("Не бывает «переименовано в семи файлах из девяти»", {"size": 11, "color": LILAC}),
])
card(s, 0.49, 3.95, 9.0, 1.15, line=PURPLE, size=12, text=[
    ("ПРОВЕРКА АКТУАЛЬНОСТИ", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 6}),
    "Код изменился между поиском и применением → рефакторинг не применяется; требуется повторный поиск",
])
notes(s, "Контраст с п. 4: там N мест – N независимых операций; здесь одна. Строка IDE: «There were changes in code after usages have been found».")

slide_plan(done={"Введение", "Подход 1: текстовый", "Подход 2: LSP", "Подход 3: движок IDE"}, current="Сравнение", expand="Подходы к рефакторингу в агентах")

# ---- Таблица ----
s = slide_blank("Сравнение", "Полнота и корректность")
cols = ["", "Текст", "LSP", "Движок IDE"]
rows = [
    ("Полнота\nнаходит все места?", [("✗", "имя ≠ символ"), ("◐", "один язык"), ("✓", "все языки")]),
    ("Корректность\nсохраняет поведение?", [("✗", "без проверки"), ("◐", "проверяет, не сообщает"), ("✓", "конфликты до правки")]),
]
x0, y0 = 0.6, 1.5
cw = [2.7, 2.1, 2.1, 2.1]
rh = 1.35
x = x0
for i, c in enumerate(cols):
    if c:
        textbox(s, x, y0, cw[i], 0.45, c, size=14, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
    x += cw[i]
hline(s, x0, y0 + 0.5, sum(cw), color=PURPLE, width=1.2)
colmap = {"✗": RED, "◐": AMBER, "✓": GREEN}
for r, (name, vals) in enumerate(rows):
    y = y0 + 0.6 + r * rh
    textbox(s, x0, y, cw[0], rh, name, size=12.5, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    x = x0 + cw[0]
    for i, (sym, lab) in enumerate(vals):
        textbox(s, x, y + 0.1, cw[i + 1], 0.6, sym, size=30, color=colmap[sym], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        textbox(s, x, y + 0.7, cw[i + 1], 0.5, lab, size=11.5, color=GREY, align=PP_ALIGN.CENTER)
        x += cw[i + 1]
    if r == 0:
        hline(s, x0, y + rh, sum(cw), color=GREY, width=0.5)
notes(s, "Не читать вслух. Устно: у текста ниша – некомпилирующийся код и нулевая цена входа. "
         "Ячейка LSP/корректность: проверка есть, но результат не проходит через протокол – см. j → i.")

slide_plan(done={"Введение", "Подходы к рефакторингу в агентах"}, current="MCP для рефакторинга")

# ---- MCP ----
s = slide_blank("Готовые решения", "MCP для рефакторинга")
textbox(s, 0.49, 1.3, 9.0, 0.6,
        "MCP (Model Context Protocol) – протокол общения LLM с инструментами. "
        "MCP-сервер – сервер, отвечающий на запросы LLM по MCP",
        size=12, color=GREY)
card(s, 0.49, 2.0, 4.4, 3.1, line=WHITE, size=11.5, gap=4, text=[
    ("ВСТРОЕННЫЙ MCP-СЕРВЕР INTELLIJ IDEA", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 8}),
    ("•", "Встроен и включён по умолчанию с версии 2025.2"),
    ("•", "Из рефакторингов – один тул: rename_refactoring"),
    ("•", "Параметры: путь в проекте, имя символа, новое имя"),
    ("•", "Extract Method, Change Signature, Inline – отсутствуют"),
])
card(s, 5.1, 2.0, 4.4, 3.1, line=WHITE, size=11.5, gap=4, text=[
    ("VEAI MCP", {"size": 10.5, "color": LILAC, "bold": True, "space_after": 8}),
    ("•", "Тулы рефакторинга IntelliJ, доступные агенту по MCP"),
    ("TODO: список тулов", {"color": AMBER}),
])
notes(s, "Встроенный сервер: адресация «путь + имя» на перегрузках неоднозначна – одна фраза. Вывод следует из сравнения карточек.")

slide_plan(done={"Введение", "Подходы к рефакторингу в агентах", "MCP для рефакторинга"}, current="Советы пользователям агентов")

# ---- Советы ----
s = slide_blank("Советы", "Советы пользователям агентов")
card(s, 0.49, 1.7, 4.4, 2.0, line=GREEN, size=12.5, gap=10, text=[
    ("Рефакторинг – отдельным коммитом", {"size": 15, "bold": True}),
    "Отдельно от изменения поведения. Ревью читаемо, откат возможен",
])
card(s, 5.1, 1.7, 4.4, 2.0, line=GREEN, size=12.5, gap=10, text=[
    ("Регрессионные тесты как фиксация поведения", {"size": 15, "bold": True}),
    "Тесты фиксируют поведение до рефакторинга; прогон после подтверждает, что поведение не изменилось",
])

# ---- Summary ----
s = slide_blank("Summary", "Главные мысли")
bullets(s, 0.6, 1.4, 8.8, 3.7, [
    "Рефакторинг – преобразование с гарантиями: полнота при поиске мест, корректность при выполнении",
    "Проблема не в модели, а в инструментах: сильная LLM с текстовыми правками ошибается там, где слабая с движком – нет",
    "Гарантии даёт только модель кода: текст не знает символов, LSP знает их в одном языке и не передаёт конфликты, движок IDE знает всё и проверяет правку до применения",
    "Модель решает «что», детерминированный инструмент – «как». Работа компилятора – не задача LLM",
    ("Движок рефакторинга IDE – готовое «как»: полнота, конфликты до правки, атомарность. Его можно отдать агенту", 0, {"color": LILAC}),
], size=13.5, gap=9)
notes(s, "По тезису – одна фраза. Пятый – последним, доклад закрывается на решении.")

# ---- Спасибо ----
s = prs.slides.add_slide(L_TEXT)
_clear_placeholders(s)
_chrome(s, TALK_SHORT)
textbox(s, 0.4, 1.6, 6.9, 2.0, ["Спасибо", "за внимание"], size=48, color=WHITE)
textbox(s, 0.4, 4.0, 6.0, 1.0, [("Вопросы?", {"size": 20, "color": LILAC, "bold": True})], size=20)
textbox(s, 5.17, 4.39, 4.42, 0.82, [
    ("Попробовать агента: veai.ru", {"bold": True}),
    ("Telegram-канал: @veai_devs", {"bold": True}),
], size=13, color=WHITE)

# ---- ЗАДЕЛ: четыре принципа (вне тайминга) ----
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
notes(s, "Backup-слайд для Q&A. Решение о включении в основной поток – после разговора с организаторами. См. «TODO внутреннее.md».")


prs.save(OUT)
print("слайдов:", len(prs.slides), "→", OUT)
