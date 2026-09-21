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
        marker = "•  " if lvl == 0 else "–  "
        indent = "      " * lvl
        r = p.add_run()
        r.text = indent + marker + it
        r.font.name = FONT
        r.font.size = Pt(st.get("size", size - 1.5 * lvl))
        r.font.color.rgb = st.get("color", color if lvl == 0 else GREY)
        r.font.bold = st.get("bold", False)
        r.font.italic = st.get("italic", False)
    return tb


def card(slide, x, y, w, h, line=WHITE, width=2.25, fill=None):
    """Скруглённая плашка в стиле шаблона: без заливки, цветная обводка."""
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
    s.text_frame.text = ""
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
    card(s, 0.5, 1.35, 9.0, 3.85, line=AMBER)
    textbox(s, 0.75, 1.5, 8.5, 0.4, "TODO: " + what, size=13, color=AMBER, bold=True)
    bullets(s, 0.75, 2.0, 8.5, 3.1, items, size=13.5, color=WHITE, gap=8)
    return s


def slide_demo(kicker, title, steps, sub=None):
    """Демо: слева область под видео, справа – шаги агента."""
    s = slide_blank(kicker, title)
    card(s, 0.49, 1.35, 5.9, 3.8, line=AMBER)
    textbox(s, 0.7, 1.45, 5.5, 0.35, "TODO: видео", size=12, color=AMBER, bold=True)
    if sub:
        textbox(s, 0.7, 1.85, 5.5, 1.2, sub, size=12, color=GREY)
    card(s, 6.6, 1.35, 2.9, 3.8, line=WHITE)
    textbox(s, 6.78, 1.45, 2.6, 0.35, "ДЕЙСТВИЯ АГЕНТА", size=10.5, color=LILAC, bold=True)
    y = 1.9
    for i, st in enumerate(steps):
        textbox(s, 6.78, y, 0.35, 0.4, str(i + 1), size=15, color=PURPLE, bold=True)
        textbox(s, 7.1, y + 0.02, 2.3, 0.6, st, size=11.5, color=WHITE)
        y += 0.62
    return s


def three_cards(s, titles, bodies, y=1.45, h=2.7, numbered=True, line=WHITE):
    xs = [0.49, 3.57, 6.65]
    for i, (t, b) in enumerate(zip(titles, bodies)):
        card(s, xs[i], y, 2.85, h, line=line if not isinstance(line, list) else line[i])
        if numbered:
            textbox(s, xs[i] + 0.3, y + 0.08, 0.6, 0.6, str(i + 1), size=34,
                    color=PURPLE, bold=True)
            textbox(s, xs[i] + 0.3, y + 0.7, 2.3, 0.6, t, size=15, color=WHITE, bold=True)
            textbox(s, xs[i] + 0.3, y + 1.3, 2.3, h - 1.4, b, size=11.5, color=WHITE)
        else:
            textbox(s, xs[i] + 0.3, y + 0.15, 2.3, 0.6, t, size=15, color=WHITE, bold=True)
            textbox(s, xs[i] + 0.3, y + 0.75, 2.3, h - 0.9, b, size=11.5, color=WHITE)


def bottom_banner(s, text, y=4.4, size=15):
    card(s, 1.2, y, 7.6, 0.68, line=PURPLE)
    textbox(s, 1.3, y + 0.05, 7.4, 0.58, text, size=size, color=WHITE, bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def two_columns(s, left_title, left_items, right_title, right_items,
                left_color=RED, right_color=GREEN, y=1.35, h=3.75,
                left_kicker=None, right_kicker=None):
    card(s, 0.49, y, 4.4, h, line=left_color)
    card(s, 5.1, y, 4.4, h, line=right_color)
    if left_kicker:
        textbox(s, 0.7, y + 0.15, 4.0, 0.3, left_kicker, size=10.5, color=left_color, bold=True)
    if right_kicker:
        textbox(s, 5.3, y + 0.15, 4.0, 0.3, right_kicker, size=10.5, color=right_color, bold=True)
    ty = y + (0.45 if left_kicker else 0.15)
    textbox(s, 0.7, ty, 4.0, 0.5, left_title, size=16, color=WHITE, bold=True)
    textbox(s, 5.3, ty, 4.0, 0.5, right_title, size=16, color=WHITE, bold=True)
    bullets(s, 0.7, ty + 0.55, 4.0, h - 1.0, left_items, size=12, gap=5)
    bullets(s, 5.3, ty + 0.55, 4.0, h - 1.0, right_items, size=12, gap=5)


# =====================================================================
# СЛАЙДЫ  (по «План подробный.md»)
# =====================================================================

def result_cards(s, missed, extra, behavior, banner):
    """Разбор результата демо: три карточки одинаковой формы."""
    three_cards(s, ["Пропущено", "Изменено лишнее", "Изменено поведение"],
                [missed, extra, behavior], line=[RED, RED, RED])
    bottom_banner(s, banner)


# ---- 1. Титул ----
s = slide_title(["Использование инструментов", "рефакторинга IDE в ИИ-агентах"],
                "Спикер: Михаил Костицын")

# ---- 2. О спикере ----
s = slide_blank(TALK_SHORT, "О себе")
card(s, 0.49, 1.4, 9.0, 2.2, line=PURPLE)
bullets(s, 0.75, 1.55, 8.5, 2.0, [
    "Ведущий разработчик Veai – ИИ-агента для разработчиков",
    "Семь лет в разработке, значительная часть – статический анализ программ и формальная верификация",
    "С 2023 года – разработка ИИ-агентов",
], size=14, gap=8)

# ---- 3. Термины: схема ----
s = slide_blank("Введение", "Ключевые термины")
# контекст слева
card(s, 0.49, 1.5, 2.9, 3.3, line=WHITE)
textbox(s, 0.65, 1.6, 2.6, 0.35, "КОНТЕКСТ", size=10.5, color=LILAC, bold=True)
textbox(s, 0.65, 1.95, 2.6, 0.4, "что модель видит", size=12, color=GREY, italic=True)
bullets(s, 0.65, 2.4, 2.6, 2.3, [
    "системный промпт", "история чата", "прочитанные файлы", "результаты вызовов инструментов",
], size=11.5, gap=4)
# модель в центре
card(s, 3.75, 2.55, 2.5, 1.2, line=PURPLE, fill=RGBColor(0x1A, 0x12, 0x33))
textbox(s, 3.8, 2.65, 2.4, 1.0, ["МОДЕЛЬ", ("генерирует ответ\nили вызов инструмента", {"size": 10.5, "color": GREY})],
        size=14, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arrow(s, 3.42, 3.08, w=0.3, h=0.14)
arrow(s, 6.28, 3.08, w=0.3, h=0.14)
# инструменты справа
card(s, 6.61, 1.5, 2.9, 3.3, line=WHITE)
textbox(s, 6.77, 1.6, 2.6, 0.35, "ИНСТРУМЕНТЫ (TOOLS, ТУЛЫ)", size=10.5, color=LILAC, bold=True)
textbox(s, 6.77, 1.95, 2.6, 0.4, "что модель может вызвать", size=12, color=GREY, italic=True)
bullets(s, 6.77, 2.4, 2.6, 2.3, [
    "чтение файла", "поиск по тексту", "правка файла", "запуск тестов",
], size=11.5, gap=4)
# возврат результата: подпись внизу
# возврат результата: стрелка снизу от инструментов к контексту
ret = s.shapes.add_shape(MSO_SHAPE.LEFT_ARROW, Inches(3.42), Inches(4.4), Inches(3.16), Inches(0.16))
ret.fill.solid(); ret.fill.fore_color.rgb = LILAC; ret.line.color.rgb = LILAC; ret.shadow.inherit = False
textbox(s, 3.42, 4.58, 3.16, 0.3, "результат вызова – в контекст", size=10.5, color=GREY, align=PP_ALIGN.CENTER)
textbox(s, 0.49, 4.95, 9.0, 0.35,
        "ИИ-агент – модель + контекст + инструменты; работает циклом",
        size=11.5, color=GREY, align=PP_ALIGN.CENTER)
notes(s, "Только три термина. LSP, PSI, MCP – по ходу. После демо – возврат: агент искал текстом и правил текстом, это были его инструменты.")

# ---- 4. Демо 1: текст ----
s = slide_demo("Демо", "Демо 1 из 3: текстовые правки",
               ["Получает задачу: переименовать метод",
                "Ищет вхождения имени по тексту",
                "Читает найденные файлы",
                "Правит каждое место текстом",
                "Отдаёт diff"],
               sub="Java/Kotlin-проект: перегрузка, метод с целевым именем, Java-переопределение, "
                   "вызов из Kotlin через свойство (user.name), строка лога, Spring XML")
notes(s, "См. demo-spec.md и «TODO внутреннее.md».")

# ---- 5. Разбор демо 1 ----
s = slide_blank("Демо", "Демо 1 из 3: результат")
result_cards(s,
             "Вызов из Kotlin через синтаксис свойства (user.name) – подстроки getName нет. Сборка не проходит",
             "Перегрузка getName(Locale) и строка лога – совпадение подстроки. Сборка проходит",
             "Вызов привязался к другому методу с тем же именем. Сборка и тесты проходят, поведение другое",
             "Проблема не в модели, а в инструментах")
notes(s, "Возврат к терминам: агент искал текстом и правил текстом. Третья карточка – сборка и тесты проходят, поведение изменилось.")

# ---- 6. Рамка ----
s = slide_blank("Идея", "Рефакторинг – преобразование с гарантиями")
card(s, 0.72, 1.35, 8.56, 0.7, line=PURPLE)
textbox(s, 0.8, 1.4, 8.4, 0.6, "Рефакторинг – изменение структуры кода без изменения поведения программы",
        size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
card(s, 0.72, 2.3, 4.1, 1.75, line=WHITE)
textbox(s, 0.95, 2.4, 3.7, 0.4, "Подзадача 1: найти все места", size=15, color=WHITE, bold=True)
textbox(s, 0.95, 2.85, 3.7, 0.8, "Нужна полнота. 99 % не годится", size=12, color=WHITE)
textbox(s, 0.95, 3.45, 3.7, 0.5, "200 вызовов × 99 % = 2 пропуска", size=12, color=LILAC, bold=True)
card(s, 5.18, 2.3, 4.1, 1.75, line=WHITE)
textbox(s, 5.41, 2.4, 3.7, 0.4, "Подзадача 2: выполнить", size=15, color=WHITE, bold=True)
textbox(s, 5.41, 2.85, 3.7, 1.1,
        "Нужна корректность. Правка не должна менять привязку вызовов, видимость, иерархию переопределений",
        size=12, color=WHITE)
bottom_banner(s, "Модель решает «что», детерминированный инструмент – «как»", y=4.3)
notes(s, "Рамка – первое из трёх повторений (6, 18). Не обсуждаем: RAG, «зачем агент, если есть Shift+F6» – Q&A.")

# =============== ЧАСТЬ 1: подходы ===============
slide_section(1, "Подходы к рефакторингу в агенте", "Текстовый подход · LSP · движок рефакторинга IDE")

# ---- Текстовый подход: поиск ----
s = slide_blank("Подход 1 из 3 · Текстовый подход", "Поиск: полнота")
textbox(s, 0.49, 1.25, 9.0, 0.3, "Как устроено: поиск по тексту, чтение файлов, правка текстом", size=11.5, color=GREY)
two_columns(s,
            "Плюсы", [
                "Не требует настройки",
                "Любой язык",
                "Работает на некомпилирующемся коде",
            ],
            "Где ломается полнота", [
                "Имя без символа: одноимённые вхождения в другом классе, в комментариях, в строках – ложные попадания",
                "Символ без имени: алиасы импорта, обращение через свойство – пропуски",
                "Шум: частотные идентификаторы (run, config, init) дают сотни вхождений",
            ],
            left_color=GREEN, right_color=RED, y=1.6, h=2.45)
card(s, 0.49, 4.2, 9.0, 0.95, line=WHITE)
textbox(s, 0.7, 4.27, 8.6, 0.45,
        [("val label = user.name   // вызов getName(), подстроки \"getName\" нет", {"font": "Menlo", "size": 12})],
        size=12, color=WHITE)
textbox(s, 0.7, 4.72, 8.6, 0.4,
        "Исследование: 71–75 % провалов текстового поиска в агенте – найденное место не выбрано из-за шума",
        size=11.5, color=GREY)
notes(s, "Где ломается полнота: пропуски (символ без имени) и ложные попадания (имя без символа). Цифра – GrepRAG, arXiv 2601.23254, сверить.")

# ---- Текстовый подход: правка ----
s = slide_blank("Подход 1 из 3 · Текстовый подход", "Правка: корректность")
three_cards(s,
            ["Скрипт / regex", "Тул «правка фрагмента»", "Перезапись файла"],
            ["Заменяет подстроку везде, где нашёл. Правит строки и комментарии; не проверяет, что заменённое – тот самый символ",
             "Удалить / добавить / заменить. Список мест – из поиска; каждое место – отдельный вызов без проверки",
             "Модель генерирует файл целиком. Помимо целевой правки может изменить соседний код"],
            y=1.4, h=2.6, numbered=False, line=[RED, RED, RED])
bottom_banner(s, "Каждая правка – независимая текстовая операция без проверки результата", y=4.3, size=13.5)
notes(s, "Где ломается корректность: правка не знает, что меняет; не проверяет привязку вызова; N мест – N независимых операций, состояние «7 из 9» возможно.")

# ---- LSP: что это ----
s = slide_blank("Подход 2 из 3 · LSP", "LSP: Language Server Protocol")
textbox(s, 0.49, 1.35, 9.0, 0.7,
        "LSP – протокол между редактором и сервером языка: отдельный процесс понимает язык, "
        "редактор или агент спрашивает его по стандартному протоколу. Как правило, один сервер – один язык",
        size=13, color=GREY)
labels = [("Редактор", "или агент"), ("Запрос", "«переименуй символ здесь»"),
          ("Сервер языка", "знает символы"), ("Правка", "список текстовых изменений; применяет клиент")]
xs = [0.72, 3.01, 5.33, 7.61]
for i, (t, sub) in enumerate(labels):
    card(s, xs[i], 2.3, 1.67, 1.0, line=WHITE)
    textbox(s, xs[i] + 0.04, 2.36, 1.59, 0.4, t, size=14, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    textbox(s, xs[i] + 0.04, 2.74, 1.59, 0.55, sub, size=10.5, color=GREY, align=PP_ALIGN.CENTER)
    if i < 3:
        arrow(s, xs[i] + 1.8, 2.73)
card(s, 0.72, 3.6, 8.56, 1.5, line=GREEN)
textbox(s, 1.0, 3.68, 8.0, 0.4, "Что даёт", size=15, color=WHITE, bold=True)
bullets(s, 1.0, 4.1, 8.0, 1.0, [
    "Переход к определению, поиск использований, переименование по символу",
    "Один протокол для всех языков: rust-analyzer, gopls, jdtls, pyright",
], size=12, gap=4)
notes(s, "Первый подход, где есть понятие символа. Деталь схемы для атомарности: сервер возвращает текстовые правки, применяет клиент.")

# ---- LSP: демо ----
s = slide_demo("Подход 2 из 3 · LSP", "Демо 2 из 3: тот же кейс через LSP",
               ["F2 на объявлении метода",
                "Редактор отправляет запрос серверу",
                "Сервер (jdtls) ищет использования по символу",
                "Сервер возвращает список правок",
                "Редактор применяет правки"],
               sub="Тот же проект. VS Code + Java-сервер (jdtls). Без агента – переименование вызывается напрямую")

# ---- LSP: разбор ----
s = slide_blank("Подход 2 из 3 · LSP", "Демо 2 из 3: результат")
result_cards(s,
             "Вызов из Kotlin через синтаксис свойства, Spring XML – сервер о них не знает",
             "Нет: перегрузка getName(Locale) и строка лога не тронуты",
             "Вызов привязался к другому методу – если сервер не сообщил о конфликте",
             "Полнота в одном языке есть; за границей языка – нет; конфликты не сообщаются")
notes(s, "Карточки заполнены по ожиданию – переписать по факту записи.")
notes(s, "Контраст с демо 1: меньше лишнего, но пропуски там же, где нет символа для этого сервера.")

# ---- LSP: пределы ----
s = slide_blank("Подход 2 из 3 · LSP", "Пределы LSP: протокол и клиент")
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
    card(s, x, y, 4.4, 1.6, line=RED)
    textbox(s, x + 0.2, y + 0.08, 4.0, 0.4, f"{i + 1}. {t}", size=13, color=WHITE, bold=True)
    textbox(s, x + 0.2, y + 0.48, 4.0, 1.1, b, size=11.5, color=WHITE)
textbox(s, 0.49, 4.75, 9.0, 0.45,
        "На некомпилирующемся коде сервер вправе отказать. Поведение зависит от реализации сервера",
        size=11.5, color=GREY)
notes(s, "Поиск у LSP – по символам, применение – текстовое, клиентом: по подзадаче 1 ближе к движку, по подзадаче 2 – к тексту. "
         "В заметках для Q&A: issue открыт с 2016, цитаты автора протокола, багрепорты 2024, «первый ответивший побеждает» в коде VS Code. "
         "См. notes-lsp-vs-psi.md, раздел 1.")

# ---- LSP: пример j → i ----
s = slide_blank("Подход 2 из 3 · LSP", "Пример: конфликт найден, но не передан")
card(s, 0.49, 1.35, 5.2, 3.75, line=WHITE)
textbox(s, 0.7, 1.45, 4.8, 0.3, "JAVA, VS CODE: ПЕРЕИМЕНОВАНИЕ j → i", size=10.5, color=GREY, bold=True)
textbox(s, 0.7, 1.9, 4.8, 2.4, [
    ("void m() {", {}),
    ("  int i = 0;", {}),
    ("  class B {", {}),
    ("    int j;              // переименуем в i", {"color": LILAC}),
    ("    void k() { j = i; }", {"color": RED}),
    ("    // после правки:  i = i", {"color": RED}),
    ("  }", {}),
    ("}", {}),
], size=12, color=WHITE, font="Menlo")
textbox(s, 0.7, 4.4, 4.8, 0.6, "Сборка и тесты проходят. Поведение изменилось", size=12, color=RED, bold=True)
card(s, 5.9, 1.35, 3.6, 3.75, line=WHITE)
textbox(s, 6.1, 1.45, 3.2, 0.3, "ЧТО ПРОИЗОШЛО", size=10.5, color=GREY, bold=True)
textbox(s, 6.1, 1.95, 3.2, 3.0,
        "Движок нашёл конфликт.\n\nПротокол его не передал – правка применена без предупреждения.\n\n"
        "В Eclipse IDE – диалог с предупреждением. Через LSP – без предупреждения",
        size=13, color=WHITE)
notes(s, "Механизм (для Q&A): движок Eclipse вычислил конфликты с уровнями; LSP-обработчик прерывается только на «фатально»; "
         "функция сборки правки принимает только набор изменений – статус передать некуда. "
         "Устно: LSP не заменяет движок, а ограничивает доступ к нему.")

# ---- Движок: заявление ----
s = slide_blank("Подход 3 из 3 · Движок IDE", "Движок рефакторинга IDE (IntelliJ)")
card(s, 0.72, 1.35, 8.56, 0.95, line=PURPLE)
textbox(s, 0.9, 1.42, 8.2, 0.85,
        "Код представлен не текстом, а связанным графом символов: символ – узел, ссылка – ребро к объявлению. "
        "«Все использования» – обход ссылок, а не поиск подстроки",
        size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
three_cards(s, ["Полнота", "Конфликты до правки", "Атомарность"],
            ["Находит все ссылки, включая другие языки и конфигурацию",
             "Проверяет, что правка не изменит поведение, и показывает список конфликтов до применения",
             "Либо всё, либо ничего. Один откат на всю операцию"],
            y=2.55, h=1.85, numbered=False, line=GREEN)
notes(s, "Рассматриваем только IntelliJ. Здесь только заявляем, механизм – на следующих слайдах.")

# ---- Движок: демо ----
s = slide_demo("Подход 3 из 3 · Движок IDE", "Демо 3 из 3: агент с тулами рефакторинга IDE",
               ["Получает ту же задачу",
                "Находит символ тулом поиска",
                "Вызывает тул переименования: preview",
                "Получает список мест и конфликт",
                "Применяет; сообщает о вхождении в строке"],
               sub="Тот же проект, тот же запрос, та же модель. Veai-агент")

# ---- Движок: разбор ----
s = slide_blank("Подход 3 из 3 · Движок IDE", "Демо 3 из 3: результат")
result_cards(s,
             "Нет",
             "Нет. Строка лога не тронута; агент сообщил о вхождении в строке",
             "Нет. Конфликт «будет вызван другой метод» показан до применения",
             "Та же модель, тот же запрос, другие тулы")
notes(s, "Карточки заполнены по ожиданию – переписать по факту записи.")
notes(s, "Не «ноль пропусков», а «ноль незамеченных пропусков»: про строку лога агент сообщил.")

# =============== ЧАСТЬ 2: движок ===============
slide_section(2, "Движок рефакторинга IDE", "Полнота · корректность · атомарность · цена")

# ---- Полнота 1: текст vs дерево ----
s = slide_blank("Полнота", "Код как текст и код как дерево")
card(s, 0.49, 1.35, 4.4, 3.35, line=RED)
textbox(s, 0.7, 1.45, 4.0, 0.3, "КАК ВИДИТ ТЕКСТОВЫЙ ПОИСК", size=10.5, color=RED, bold=True)
textbox(s, 0.7, 1.85, 4.0, 2.8, [
    ("class User {", {}),
    ("  String getName() {...}", {"color": LILAC}),
    ("  String getName(Locale l) {...}", {"color": LILAC}),
    ("}", {}),
    ("user.getName();", {"color": LILAC}),
    ("log.info(\"getName called\");", {"color": LILAC}),
    ("order.getName();", {"color": LILAC}),
    ("val label = user.name  // Kotlin", {"color": GREY}),
], size=11.5, color=WHITE, font="Menlo")
card(s, 5.1, 1.35, 4.4, 3.35, line=GREEN)
textbox(s, 5.3, 1.45, 4.0, 0.3, "КАК ВИДИТ IDE", size=10.5, color=GREEN, bold=True)
textbox(s, 5.3, 1.85, 4.0, 2.8, [
    ("User.getName()       объявление", {"color": GREEN}),
    ("User.getName(Locale) другой символ", {"color": GREY}),
    ("user.getName()  ──▶  ссылка", {"color": GREEN}),
    ("\"getName called\"     строка, не ссылка", {"color": GREY}),
    ("order.getName() ──▶  Order.getName()", {"color": GREY}),
    ("user.name       ──▶  ссылка (Kotlin)", {"color": GREEN}),
], size=11.5, color=WHITE, font="Menlo")
textbox(s, 0.49, 4.8, 9.0, 0.4,
        "PSI (Program Structure Interface) – модель кода: дерево + ссылки между узлами. IDE держит её в памяти и синхронизирует с редактором",
        size=11.5, color=LILAC, align=PP_ALIGN.CENTER)
notes(s, "Вводим PSI. Слева нет понятия символа, справа каждый вызов знает своё объявление. Дальше – «модель кода».")

# ---- Полнота 2: ссылки ----
s = slide_blank("Полнота", "Откуда движок знает все места")
card(s, 0.72, 2.0, 2.2, 0.8, line=PURPLE)
textbox(s, 0.75, 2.05, 2.14, 0.7, ["getName()", ("объявление", {"size": 9.5, "color": GREY})],
        size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
refs = [("вызов в Java", 3.6, 1.5), ("вызов из Kotlin: user.name", 3.6, 2.1), ("имя свойства в Spring XML", 3.6, 2.7)]
for t, x, y in refs:
    card(s, x, y, 2.2, 0.5, line=WHITE)
    textbox(s, x + 0.03, y + 0.06, 2.14, 0.4, t, size=11, color=WHITE, align=PP_ALIGN.CENTER)
    arrow(s, 3.05, y + 0.18, w=0.5, h=0.14)
textbox(s, 0.72, 2.85, 2.2, 0.6, "ссылка: источник → цель;\nобратный поиск – все источники", size=10.5, color=GREY, align=PP_ALIGN.CENTER)
card(s, 6.2, 1.4, 3.3, 2.3, line=GREEN)
textbox(s, 6.4, 1.47, 3.0, 0.3, "ИСТОЧНИКИ ССЫЛОК", size=10.5, color=GREEN, bold=True)
bullets(s, 6.4, 1.8, 3.0, 1.9, [
    "Ссылка – слой поверх дерева",
    "Плагины добавляют ссылки в другие языки и конфигурацию",
    "Индексы: где какое имя встречается, кто кого наследует",
    "Индексы и модель – в памяти IDE и кэшах на диске",
], size=11, gap=3)
textbox(s, 0.72, 3.5, 5.3, 0.5,
        "На некомпилирующемся коде – частично: неразрешённые ссылки помечаются отдельно",
        size=11, color=GREY)
card(s, 0.72, 4.05, 8.78, 1.05, line=RED)
textbox(s, 0.95, 4.12, 8.4, 0.3, "ОГРАНИЧЕНИЕ", size=10.5, color=RED, bold=True)
textbox(s, 0.95, 4.4, 8.4, 0.7, [
    ("Class.forName(\"com.foo.Bar\")   – ссылка есть, переименуется", {}),
    ("Class.forName(prefix + name)   – runtime-рефлексия, не анализируется", {}),
], size=11.5, color=WHITE, font="Menlo")
notes(s, "Ссылка – слой поверх дерева, поэтому плагин может добавить связь туда, где компилятор её не видит. Индексы – почему обход быстрый; цена – п. 11. Граница – сказать самому.")

# ---- Корректность ----
s = slide_blank("Корректность", "Откуда движок знает, что не сломает")
card(s, 0.72, 1.3, 8.56, 0.6, line=PURPLE)
textbox(s, 0.8, 1.33, 8.4, 0.55, "Проверки «имя занято» недостаточно – движок проверяет правку до применения",
        size=13.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
steps = [("Копия вызова", "оригинал не трогаем"), ("Новое имя", "на копии"),
         ("Разрешение", "заново: какой метод?"), ("Другой метод", "→ конфликт")]
xs = [0.72, 3.01, 5.33, 7.61]
for i, (t, sub) in enumerate(steps):
    card(s, xs[i], 2.1, 1.67, 0.8, line=WHITE)
    textbox(s, xs[i] + 0.04, 2.14, 1.59, 0.4, t, size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    textbox(s, xs[i] + 0.04, 2.5, 1.59, 0.35, sub, size=10.5, color=GREY, align=PP_ALIGN.CENTER)
    if i < 3:
        arrow(s, xs[i] + 1.8, 2.43)
# пример из демо
card(s, 0.72, 3.1, 3.7, 1.95, line=WHITE)
textbox(s, 0.9, 3.15, 3.4, 0.3, "ДО", size=10.5, color=GREY, bold=True)
textbox(s, 0.9, 3.45, 3.4, 0.5, [("account.getName(locale)", {"font": "Menlo"})], size=12, color=WHITE)
textbox(s, 0.9, 3.95, 3.4, 0.9, "привязан к getName(Object) – его переименуем", size=11.5, color=GREEN)
card(s, 5.58, 3.1, 3.7, 1.95, line=WHITE)
textbox(s, 5.76, 3.15, 3.4, 0.3, "ПОСЛЕ", size=10.5, color=GREY, bold=True)
textbox(s, 5.76, 3.45, 3.4, 0.5, [("account.getDisplayName(locale)", {"font": "Menlo"})], size=12, color=WHITE)
textbox(s, 5.76, 3.95, 3.4, 0.9, "привязался бы к существующему getDisplayName(Locale) – другой метод", size=11.5, color=RED)
arrow(s, 4.7, 3.98, w=0.6, h=0.16)
textbox(s, 0.72, 5.05, 8.56, 0.3, "Конфликт: «После переименования будет вызван другой метод»",
        size=11.5, color=LILAC, bold=True, align=PP_ALIGN.CENTER)
notes(s, "Ловушка 2 из demo-spec.md: текст применил молча, LSP применил молча, движок остановился. Существующий метод более специфичен, чем переименуемый – поэтому вызов пересаживается. Механизм – RenameJavaMethodProcessor, advancedResolve.")

# ---- Атомарность ----
s = slide_blank("Атомарность", "Почему не бывает «наполовину»")
card(s, 0.49, 1.4, 4.4, 2.3, line=WHITE)
textbox(s, 0.7, 1.5, 4.0, 0.4, "Одна команда", size=15, color=WHITE, bold=True)
textbox(s, 0.7, 1.95, 4.0, 1.0, "Поиск использований, проверка конфликтов и правка – одна команда движка", size=12, color=WHITE)
textbox(s, 0.7, 2.95, 4.0, 0.6, "Как транзакция в БД: либо все правки, либо ни одной", size=11, color=LILAC, italic=True)
card(s, 5.1, 1.4, 4.4, 2.3, line=WHITE)
textbox(s, 5.3, 1.5, 4.0, 0.4, "Один откат", size=15, color=WHITE, bold=True)
textbox(s, 5.3, 1.95, 4.0, 1.0, "Глобальный откат на всю операцию, а не по файлу. Одно действие отмены", size=12, color=WHITE)
textbox(s, 5.3, 2.95, 4.0, 0.6, "Не бывает «переименовано в семи файлах из девяти»", size=11, color=LILAC, italic=True)
card(s, 0.49, 3.95, 9.0, 1.15, line=PURPLE)
textbox(s, 0.7, 4.02, 8.6, 0.3, "ПРОВЕРКА АКТУАЛЬНОСТИ", size=10.5, color=LILAC, bold=True)
textbox(s, 0.7, 4.32, 8.6, 0.7, "Код изменился между поиском и применением → рефакторинг не применяется; требуется повторный поиск",
        size=12, color=WHITE)
notes(s, "Контраст с п. 4: там N мест – N независимых операций; здесь одна. Строка IDE: «There were changes in code after usages have been found».")

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

# ---- Цена ----
s = slide_blank("Ограничения", "Цена подхода")
card(s, 0.49, 1.35, 9.0, 3.3, line=WHITE)
bullets(s, 0.75, 1.55, 8.5, 3.0, [
    "Нужна запущенная IDE с индексами и кэшами",
    "Индексация занимает время; в этот период рефакторинг недоступен",
    "Применение правки – под блокировкой на запись: редактор ждёт",
    "Несколько копий проекта (worktree) – отдельная IDE и индексы на каждую",
    "CI и удалённые машины – стандартного способа нет; официальный CLI рефакторингов не поддерживает",
], size=13.5, gap=10)
textbox(s, 0.49, 4.75, 9.0, 0.35, "TODO: worktree и CI – уточнить формулировки по практике", size=11, color=AMBER)
notes(s, "Тон – список фактов, без «но зато». Строка IDE: «Safe delete refactoring is not available during the project analysis».")

# ---- Границы ----
s = slide_blank("Ограничения", "Границы подхода")
card(s, 0.49, 1.35, 9.0, 3.0, line=RED)
textbox(s, 0.75, 1.45, 8.5, 0.35, "ЧТО ДВИЖОК НЕ ДЕЛАЕТ", size=10.5, color=RED, bold=True)
bullets(s, 0.75, 1.9, 8.5, 2.4, [
    "Runtime-рефлексия, имена через конкатенацию – не анализируются",
    "Связи через SQL, шаблоны, HTTP-контракты – вне модели кода",
    "Архитектурные изменения (разделить сервис, поменять модель данных) – не рефакторинг в смысле движка",
], size=13.5, gap=10)
notes(s, "Движок гарантирует полноту и корректность внутри того, что видит как ссылки. Всё, что связано не ссылкой, – вне.")

# ---- MCP ----
s = slide_blank("Готовые решения", "MCP для рефакторинга")
textbox(s, 0.49, 1.3, 9.0, 0.6,
        "MCP (Model Context Protocol) – протокол общения LLM с инструментами. "
        "MCP-сервер – сервер, отвечающий на запросы LLM по MCP",
        size=12, color=GREY)
card(s, 0.49, 2.0, 4.4, 3.1, line=WHITE)
textbox(s, 0.7, 2.1, 4.0, 0.3, "ВСТРОЕННЫЙ MCP-СЕРВЕР INTELLIJ IDEA", size=10.5, color=LILAC, bold=True)
bullets(s, 0.7, 2.45, 4.0, 2.6, [
    "Встроен и включён по умолчанию с версии 2025.2",
    "Из рефакторингов – один тул: rename_refactoring",
    ("Параметры: путь в проекте, имя символа, новое имя", 1),
    "Extract Method, Change Signature, Inline – отсутствуют",
], size=11.5, gap=4)
card(s, 5.1, 2.0, 4.4, 3.1, line=WHITE)
textbox(s, 5.3, 2.1, 4.0, 0.3, "VEAI MCP", size=10.5, color=LILAC, bold=True)
bullets(s, 5.3, 2.45, 4.0, 2.6, [
    "Тулы рефакторинга IntelliJ, доступные агенту по MCP",
    ("TODO: список тулов", 0, {"color": AMBER}),
], size=11.5, gap=4)
notes(s, "Встроенный сервер: адресация «путь + имя» на перегрузках неоднозначна – одна фраза. Вывод следует из сравнения карточек.")

# ---- Советы ----
s = slide_blank("Советы", "Советы пользователям агентов")
card(s, 0.49, 1.7, 4.4, 2.0, line=GREEN)
textbox(s, 0.7, 1.85, 4.0, 0.6, "Рефакторинг – отдельным коммитом", size=15, color=WHITE, bold=True)
textbox(s, 0.7, 2.55, 4.0, 1.0, "Отдельно от изменения поведения. Ревью читаемо, откат возможен", size=12.5, color=WHITE)
card(s, 5.1, 1.7, 4.4, 2.0, line=GREEN)
textbox(s, 5.3, 1.85, 4.0, 0.6, "Регрессионные тесты как фиксация поведения", size=15, color=WHITE, bold=True)
textbox(s, 5.3, 2.55, 4.0, 1.0, "Тесты фиксируют поведение до рефакторинга; прогон после подтверждает, что поведение не изменилось", size=12.5, color=WHITE)

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
    card(s, xs4[i], 1.4, 2.15, 3.7, line=WHITE)
    textbox(s, xs4[i] + 0.15, 1.5, 1.85, 0.5, t, size=13.5, color=WHITE, bold=True)
    textbox(s, xs4[i] + 0.15, 2.05, 1.85, 3.0, b, size=11.5, color=WHITE)
notes(s, "Backup-слайд для Q&A. Решение о включении в основной поток – после разговора с организаторами. См. «TODO внутреннее.md».")


prs.save(OUT)
print("слайдов:", len(prs.slides), "→", OUT)
