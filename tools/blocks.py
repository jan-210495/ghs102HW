# -*- coding: utf-8 -*-
"""وحدات بناء بسيطة لوصف بنية الوظيفة الأكاديمية.

يُستعمل هذا الملف بواسطة ملفات المحتوى وأداة البناء؛ لا يحتوي على نص البحث نفسه.
"""


def _b(t, **kw):
    d = {"t": t}
    d.update(kw)
    return d


def cover_title(text, size=None):
    return _b("cover_title", text=text, size=size)


def center(text, bold=False, size=None, color=None, space_after=6):
    return _b("center", text=text, bold=bold, size=size, color=color, space_after=space_after)


def h1(text, page=None):
    return _b("h1", text=text, page=page)


def h2(text, page=None):
    return _b("h2", text=text, page=page)


def h3(text):
    return _b("h3", text=text)


def p(text, indent=True):
    return _b("p", text=text, indent=indent)


def bullets(items):
    return _b("bullets", items=items)


def numbers(items):
    return _b("numbers", items=items)


def table(headers, rows, widths=None, caption=None):
    return _b("table", headers=headers, rows=rows, widths=widths, caption=caption)


def note(text, label="ملاحظة منهجية"):
    return _b("note", text=text, label=label)


def quote(text, source=None):
    return _b("quote", text=text, source=source)


def pb():
    return _b("pb")


def hr():
    return _b("hr")


def spacer(pts=6):
    return _b("spacer", pts=pts)


def signature(lines):
    return _b("signature", lines=lines)


def toc(rows):
    return _b("toc", rows=rows)


def refs_block():
    return _b("refs")


def wordcount_block():
    return _b("wordcount")


def toc_field():
    return _b("toc_field")
