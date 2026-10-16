"""Перенести комментарии PowerPoint (modern comments) из рабочей презентации в свежую сборку.

Запуск:  python3 merge_comments.py "<рабочий .pptx>" "<собранный .pptx>"
Комментарии привязываются по номеру слайда (позиция в списке), поэтому порядок слайдов в обеих
презентациях должен совпадать. Сборка build_deck.py комментарии не создаёт, а Михаил пишет в них
свои TODO – без переноса они терялись бы при каждой пересборке.
"""
import re
import shutil
import sys
import tempfile
import zipfile

CT_COMMENTS = "application/vnd.ms-powerpoint.comments+xml"
CT_AUTHORS = "application/vnd.ms-powerpoint.authors+xml"
REL_COMMENTS = "http://schemas.microsoft.com/office/2018/10/relationships/comments"
REL_AUTHORS = "http://schemas.microsoft.com/office/2018/10/relationships/authors"


def slide_order(z):
    """Список имён частей слайдов в порядке показа (по presentation.xml и его rels)."""
    pres = z.read("ppt/presentation.xml").decode()
    rels = z.read("ppt/_rels/presentation.xml.rels").decode()
    targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    targets.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    rids = re.findall(r'<p:sldId [^>]*r:id="([^"]+)"', pres)
    return ["ppt/" + targets[r].lstrip("/").replace("ppt/", "", 1) if not targets[r].startswith("ppt/") else targets[r] for r in rids]


def main(src, dst, new_slides=()):
    """new_slides – номера слайдов (с 1), которых в собранном файле нет в рабочем: комментарии сдвигаются мимо них."""
    with zipfile.ZipFile(src) as zs:
        comments = {}
        order = slide_order(zs)
        for i, part in enumerate(order):
            rel_name = part.replace("slides/", "slides/_rels/") + ".rels"
            if rel_name not in zs.namelist():
                continue
            rels = zs.read(rel_name).decode()
            m = re.search(r'Target="\.\./comments/([^"]+)"', rels)
            if m:
                comments[i] = ("ppt/comments/" + m.group(1), zs.read("ppt/comments/" + m.group(1)))
        authors = zs.read("ppt/authors.xml") if "ppt/authors.xml" in zs.namelist() else None
    if not comments:
        print("комментариев нет – ничего не делаю")
        return
    tmp = tempfile.mktemp(suffix=".pptx")
    with zipfile.ZipFile(dst) as zd, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        order = slide_order(zd)
        skip = {n - 1 for n in new_slides}
        dst_idx = [i for i in range(len(order)) if i not in skip]       # рабочий слайд k → собранный dst_idx[k]
        comments = {dst_idx[i]: v for i, v in comments.items() if i < len(dst_idx)}
        names = set(zd.namelist())
        for item in zd.infolist():
            data = zd.read(item.filename)
            if item.filename == "[Content_Types].xml":
                s = data.decode()
                add = "".join(f'<Override PartName="/{p}" ContentType="{CT_COMMENTS}"/>' for p, _ in comments.values())
                if authors is not None and "/ppt/authors.xml" not in s:
                    add += f'<Override PartName="/ppt/authors.xml" ContentType="{CT_AUTHORS}"/>'
                data = s.replace("</Types>", add + "</Types>").encode()
            elif item.filename == "ppt/_rels/presentation.xml.rels" and authors is not None and "authors.xml" not in data.decode():
                data = data.decode().replace("</Relationships>", f'<Relationship Id="rIdAuthors" Type="{REL_AUTHORS}" Target="authors.xml"/></Relationships>').encode()
            else:
                for i, (part, _) in comments.items():
                    rel_name = order[i].replace("slides/", "slides/_rels/") + ".rels"
                    if item.filename == rel_name:
                        data = data.decode().replace("</Relationships>", f'<Relationship Id="rIdComments" Type="{REL_COMMENTS}" Target="../comments/{part.split("/")[-1]}"/></Relationships>').encode()
            zo.writestr(item, data)
        for part, blob in comments.values():
            zo.writestr(part, blob)
        if authors is not None and "ppt/authors.xml" not in names:
            zo.writestr("ppt/authors.xml", authors)
    shutil.move(tmp, dst)
    print(f"перенесено комментариев: {len(comments)} (слайды {', '.join(str(i + 1) for i in sorted(comments))})")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], [int(n) for n in sys.argv[3].split(",")] if len(sys.argv) > 3 else ())
