#!/usr/bin/env python3
"""Перемонтаж трёх демо в ролики по ≈ 60 с по плану CUT-PLAN-60s.md.

    python3 cut60.py            # все три: runs/<run>/session-60s.mp4 + captions-60s.ass
    python3 cut60.py 3          # только демо 3
    python3 cut60.py --ass-only # только оверлеи подписей и таблица таймкодов

Без титулов и финальных карточек: только экран агента и подписи на тёмной полосе внизу
(полоса закрывает строку ввода, а не вывод агента). Исходники не меняются.
Таймкоды сегментов – в секундах исходного session.mp4 (как в STORYLINES.md).
"""
import os, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
FONTS = os.path.join(HERE, "fonts")
BG = "0x191A23"
BAND_H = 150

# (начало, конец, скорость, подпись или None); ("hold", t, длительность, подпись) – стоп-кадр
RUNS = {
    "1": dict(
        dir="runs/1-claude-code", crop=None,
        segs=[
            (11, 20, 3, r"Инструменты агента: grep, чтение файлов, скрипты"),
            (59, 68, 1, r"Скрипт дописывает compile = false к старым вызовам.\NУпал – починил через sed"),
            (76, 82, 2, r"Слияние функций – точечными правками"),
            (83, 91, 1, r"Переименование по подстроке в 131 файле –\Nи в комментарии тоже: «not lint for performance reasons»"),
            (110, 144, 5, r"Тесты запущены без флага -Pcompile-test-snippets=true"),
            ("hold", 386, 6, r"Правка попала в соседний модуль detekt-sample-extensions:\Nон собирается против 1.23.7, где параметра compile нет"),
        ]),
    "2": dict(
        dir="runs/2-claude-code-serena", crop=None,
        segs=[
            (10, 18, 3, r"Serena – инструменты LSP по MCP"),
            (41, 47, 1, r"Агент сам объясняет, зачем LSP:\Ngrep не отличит lint(KtFile) от lint(String)"),
            (47, 62, 4, r"rename_symbol: старый lint → временное имя, по ссылкам"),
            (62, 68, 1, r"Сервер: «24 changes applied». Вызовы без точки найдены,\Nодноимённые функции не тронуты"),
            (68, 77, 1, r"Сервер заодно удалил посторонний импорт и не сообщил.\NАгент заметил по диффу и откатил"),
            (106, 113, 1, r"Добавить параметр – смена сигнатуры, в протоколе её нет → скрипт.\N18 файлов из 116, починил через sed"),
            (122, 157, 7, r"Слияние функций и комментарии – текстовыми заменами"),
            (179, 236, 15, r"Тесты с флагом -Pcompile-test-snippets=true"),
            ("hold", 275, 4, None),
        ]),
    "3": dict(
        dir="runs/3-veai", crop="crop=1360:1080:500:112,pad=1920:1080:280:0:color=" + BG,
        segs=[
            (20, 26, 1, r"find_usages: 99 мест – по ссылкам, а не по тексту"),
            (40, 52, 1.5, r"Rename движком: пробный прогон – 23 файла, 0 конфликтов.\NИмя lint свободно"),
            (52, 62, 1.5, r"Второй Rename: 130 файлов, 0 конфликтов"),
            (62, 86, 8, r"Задвоенные импорты в 18 файлах движок схлопнул сам"),
            (113, 125, 3, r"Модель решает «что», движок – «как»"),
            (150, 158, 1, r"Рефакторинга «добавить аргумент во все вызовы» нет →\Nperl по временному имени: в нём только старые вызовы"),
            (214, 247, 5, r"Сборка в терминале упала (JDK 25) – агент запускает через раннер IDE"),
            (262, 268, 2, r"18 тестов прошли"),
            (275, 296, 4, r"apiCheck упал на пустой строке – агент убрал её"),
            ("hold", 300, 4, None),
        ]),
}

from PIL import Image, ImageDraw, ImageFont
FONT = os.path.join(FONTS, "GolosText.ttf")


def caption_png(path, text, size=46):
    """Прозрачный кадр 1920×1080 с подписью по центру нижней полосы (сама полоса рисуется ffmpeg)."""
    im = Image.new("RGBA", (1920, 1080), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    f = ImageFont.truetype(FONT, size)
    lines = text.split("\\N")
    lh = size + 10
    y = 1080 - BAND_H + (BAND_H - lh * len(lines)) / 2
    for ln in lines:
        w = d.textlength(ln, font=f)
        d.text(((1920 - w) / 2, y), ln, font=f, fill=(255, 255, 255, 255))
        y += lh
    im.save(path)


def badge_png(path, text):
    im = Image.new("RGBA", (1920, 1080), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    f = ImageFont.truetype(FONT, 30)
    w = d.textlength(text, font=f)
    x1, y1 = 1920 - 40 - w - 24, 40
    d.rounded_rectangle((x1, y1, 1920 - 40, y1 + 50), radius=8, fill=(0x6A, 0x5C, 0xFF, 200))
    d.text((x1 + 12, y1 + 8), text, font=f, fill=(255, 255, 255, 255))
    im.save(path)


def ts(x):
    cs = int(round(x * 100))
    h, cs = divmod(cs, 360000); m, cs = divmod(cs, 6000); s, cs = divmod(cs, 100)
    return f"{h}:{m:02d}:{s:02d}.{cs:02d}"


def seg_len(s):
    return s[2] if s[0] == "hold" else (s[1] - s[0]) / s[2]


def build(key, ass_only):
    r = RUNS[key]
    d = os.path.join(HERE, r["dir"])
    src, out = os.path.join(d, "session.mp4"), os.path.join(d, "session-60s.mp4")
    ovdir = os.path.join(d, "overlays-60s"); os.makedirs(ovdir, exist_ok=True)
    starts, t = [], 0.0
    for s in r["segs"]:
        starts.append(t); t += seg_len(s)
    overlays = []   # (png, a, b)
    badges = {}
    for i, s in enumerate(r["segs"]):
        a, b = starts[i], starts[i] + seg_len(s)
        if s[3]:
            png = os.path.join(ovdir, f"cap{i:02d}.png"); caption_png(png, s[3]); overlays.append((png, a, b - 0.1))
        if s[0] != "hold" and s[2] != 1:
            sp = f"{s[2]:g}"
            if sp not in badges:
                badges[sp] = os.path.join(ovdir, f"fast{sp}.png"); badge_png(badges[sp], f"ускорено ×{sp}")
            overlays.append((badges[sp], a, b))
    print(f"== демо {key}: {ts(t)}")
    for i, s in enumerate(r["segs"]):
        srcspan = f"кадр {s[1]}" if s[0] == "hold" else f"{s[0]}–{s[1]} ×{s[2]:g}"
        print(f"  {ts(starts[i])}  {srcspan:18} {(s[3] or '').replace(chr(92) + 'N', ' / ')}")
    if ass_only:
        return
    pre = (r["crop"] + ",") if r["crop"] else ""
    args, chains, labels, n = ["ffmpeg", "-hide_banner", "-loglevel", "error", "-y"], [], [], 0
    for i, s in enumerate(r["segs"]):
        if s[0] == "hold":
            args += ["-ss", str(s[1]), "-t", "0.2", "-i", src]
            chains.append(f"[{n}:v]{pre}fps=30,setsar=1,tpad=stop_mode=clone:stop_duration={s[2]},trim=duration={s[2]},setpts=PTS-STARTPTS,"
                          f"drawbox=x=0:y=ih-{BAND_H}:w=iw:h={BAND_H}:color={BG}@0.92:t=fill[s{i}]")
        else:
            a, b, sp = s[0], s[1], s[2]
            args += ["-ss", str(a), "-t", str(b - a), "-i", src]
            chains.append(f"[{n}:v]setpts=(PTS-STARTPTS)/{sp},fps=30,{pre}setsar=1,"
                          f"drawbox=x=0:y=ih-{BAND_H}:w=iw:h={BAND_H}:color={BG}@0.92:t=fill,"
                          f"trim=duration={(b - a) / sp:.3f},setpts=PTS-STARTPTS[s{i}]")
        n += 1; labels.append(f"[s{i}]")
    fc = ";".join(chains) + ";" + "".join(labels) + f"concat=n={len(labels)}:v=1:a=0[v0]"
    cur = "v0"
    for j, (png, a, b) in enumerate(overlays):
        args += ["-i", png]
        fc += f";[{cur}][{n}:v]overlay=0:0:enable='between(t,{a:.2f},{b:.2f})'[v{j + 1}]"
        cur = f"v{j + 1}"; n += 1
    fc += f";[{cur}]format=yuv420p[out]"
    tmp = out + ".tmp.mp4"
    args += ["-filter_complex", fc, "-map", "[out]", "-an", "-r", "30", "-c:v", "libx264", "-preset", "medium", "-crf", "18",
             "-pix_fmt", "yuv420p", "-movflags", "+faststart", tmp]
    subprocess.run(args, check=True, cwd=d)
    subprocess.run(["ffmpeg", "-v", "error", "-i", tmp, "-f", "null", "-"], check=True)
    os.replace(tmp, out)
    print(f"  → {out}")


if __name__ == "__main__":
    keys = [a for a in sys.argv[1:] if a in RUNS] or list(RUNS)
    for k in keys:
        build(k, "--ass-only" in sys.argv)
