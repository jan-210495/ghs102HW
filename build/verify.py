# -*- coding: utf-8 -*-
"""فحص الملف الناتج: سلامة البنية، والعدّ، والأصالة، واللغة."""
import json
import re
import zipfile
import pathlib
import sys
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
DOCX = pathlib.Path('الوظيفة_الأولى_الرأي_العام_جان_وديع.docx')
report = {}

PPR_SEQ = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
           'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
           'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
           'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
           'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
           'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
           'sectPr', 'pPrChange']
RPR_SEQ = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
           'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
           'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs',
           'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs',
           'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']
TBLPR_SEQ = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
             'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders',
             'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription']
TRPR_SEQ = ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit',
            'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden']

z = zipfile.ZipFile(DOCX)
parts = z.namelist()
xml_ok = True
trees = {}
for name in parts:
    if name.endswith('.xml') or name.endswith('.rels'):
        try:
            trees[name] = etree.fromstring(z.read(name))
        except Exception as e:
            xml_ok = False
            print('XML ERROR', name, e)
report['all_xml_wellformed'] = xml_ok

doc = trees['word/document.xml']


def order_errors(tree, seq, tag):
    errs = 0
    for el in tree.iter(W + tag):
        idxs = []
        for child in el:
            nm = etree.QName(child).localname
            if nm in seq:
                idxs.append(seq.index(nm))
        if idxs != sorted(idxs):
            errs += 1
    return errs


report['order_errors'] = {
    'pPr': order_errors(doc, PPR_SEQ, 'pPr'),
    'rPr': order_errors(doc, RPR_SEQ, 'rPr'),
    'tblPr': order_errors(doc, TBLPR_SEQ, 'tblPr'),
    'trPr': order_errors(doc, TRPR_SEQ, 'trPr'),
}

# ---------------------------------------------------------------- الهوامش
fn = trees.get('word/footnotes.xml')
def fn_texts(tree):
    out = []
    for f in tree.iter(W + 'footnote'):
        if f.get(W + 'type'):
            continue
        txt = ''.join(t.text or '' for t in f.iter(W + 't'))
        out.append(txt.strip())
    return out
footnotes = fn_texts(fn) if fn is not None else []
report['footnote_count'] = len(footnotes)
refs = [r.get(W + 'id') for r in doc.iter(W + 'footnoteReference')]
report['footnote_reference_count'] = len(refs)
report['footnote_ids_match'] = sorted(map(int, refs)) == list(range(1, len(footnotes) + 1))
report['empty_footnotes'] = [i + 1 for i, t in enumerate(footnotes) if not t]

# ---------------------------------------------------------------- الجداول والأشكال
report['table_count'] = len(list(doc.iter(W + 'tbl')))
report['image_count'] = len(list(doc.iter(W + 'drawing')))
report['toc_field'] = any('TOC' in (i.text or '') for i in doc.iter(W + 'instrText'))
report['pageref_fields'] = sum(1 for i in doc.iter(W + 'instrText') if 'PAGEREF' in (i.text or ''))
bms = {b.get(W + 'name') for b in doc.iter(W + 'bookmarkStart')}
targets = {re.search(r'PAGEREF\s+(\S+)', i.text).group(1)
           for i in doc.iter(W + 'instrText') if 'PAGEREF' in (i.text or '')}
report['bookmarks'] = sorted(bms - {'_GoBack'})
report['pageref_missing_targets'] = sorted(targets - bms)
report['updateFields'] = trees['word/settings.xml'].find(f'.//{W}updateFields') is not None
report['bidi_sections'] = len(list(doc.iter(W + 'bidi')))
report['pgNumType'] = [p.get(W + 'start') for p in trees['word/document.xml'].iter(W + 'pgNumType')]
report['headers'] = [n for n in parts if re.match(r'word/header\d+\.xml', n)]
report['footers'] = [n for n in parts if re.match(r'word/footer\d+\.xml', n)]

# ---------------------------------------------------------------- النصوص
def para_text(p):
    return ''.join(t.text or '' for t in p.iter(W + 't'))


def norm(t):
    t = re.sub(r'[\u064B-\u0652\u0670\u06D6-\u06ED]', '', t)
    t = re.sub(r'[«»"\(\)\[\]،؛:\.\!\؟\-\u2013\u2014]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()


def words(t):
    return [w for w in re.split(r'\s+', norm(t)) if w and not re.fullmatch(r'[\d\W]+', w)]


body = []           # (heading_level, heading_text, [paragraphs])
cur = {'level': 0, 'text': 'المقدمة', 'paras': []}
blocks = [cur]
for el in doc.find(W + 'body'):
    if el.tag == W + 'p':
        style = el.find(f'{W}pPr/{W}pStyle')
        st = style.get(W + 'val') if style is not None else ''
        txt = para_text(el)
        if st.startswith('Heading'):
            lvl = int(st[-1])
            if lvl == 1:
                cur = {'level': 1, 'text': txt, 'paras': []}
                blocks.append(cur)
            elif lvl == 2:
                cur = {'level': 2, 'text': txt, 'paras': []}
                blocks.append(cur)
        else:
            if txt.strip():
                cur['paras'].append(txt)
    elif el.tag == W + 'tbl':
        txt = ' '.join(para_text(p) for p in el.iter(W + 'p'))
        cur['paras'].append(txt)

# كلمات لكل قسم رئيسي (نص القسم + أقسامه الفرعية)
section_words = {}
current = None
for b in blocks:
    n = sum(len(words(p)) for p in b['paras'])
    if b['level'] == 1:
        current = b['text']
        section_words[current] = section_words.get(current, 0) + n
    elif current:
        section_words[current] += n
report['words_by_h1'] = section_words

def sub_words(prefix):
    tot = 0
    for b in blocks:
        if b['level'] == 1:
            cur_l = b
        elif b['level'] == 2:
            if cur_l['text'].startswith(prefix):
                tot += sum(len(words(p)) for p in b['paras'])
    return tot


q_text = {b['text']: sum(len(words(p)) for p in b['paras']) for b in blocks if b['level'] == 2}
jan = sum(v for k, v in q_text.items() if k.startswith(('1.', '2.', '3.')))
wadee = sum(v for k, v in q_text.items() if k.startswith(('4.', '5.')))
report['words_jan_questions'] = jan
report['words_wadee_questions'] = wadee
# كلمات كل باحث في الخاتمة (كلمته الأخيرة + صعوباته + ما أضافه العمل)
jan += sum(len(words(p)) for b in blocks if b['level'] == 2 and b['text'].startswith('كلمة جان') for p in b['paras'])
wadee += sum(len(words(p)) for b in blocks if b['level'] == 2 and b['text'].startswith('كلمة وديع') for p in b['paras'])
# كلمات كل باحث في الخاتمة (الكلمتان الأخيرتان + الصعوبات)
jan_words_leads = sum(len(words(p)) for p in blocks if False)  # placeholder
report['words_jan_sections'] = jan
report['words_wadee_sections'] = wadee
report['words_subsections'] = q_text

# كلمات المتن كاملاً دون الجداول والمراجع
body_words = 0
for b in blocks:
    if b['level'] in (1, 2):
        body_words += sum(len(words(p)) for p in b['paras'])
report['words_body_total'] = body_words
report['words_footnotes'] = sum(len(words(t)) for t in footnotes)

# ---------------------------------------------------------------- اللغة
all_text = '\n'.join(para_text(p) for p in doc.iter(W + 'p'))
latin = re.findall(r'[A-Za-z][A-Za-z\._\-]{1,}', all_text)
report['latin_tokens'] = sorted(set(latin))
report['latin_count'] = len(latin)

# ---------------------------------------------------------------- الأصالة: تشابه الجمل
def sentences(txt):
    out = []
    for s in re.split(r'[\.\؟\!\n]', norm(txt)):
        ws = s.split()
        if len(ws) >= 6:
            out.append(' '.join(ws))
    return out

jan_text = ' '.join(p for b in blocks if b['level'] == 2 and b['text'].startswith(('المبحث الأول', 'المبحث الثاني', 'المبحث الثالث')) for p in b['paras'])
wadee_text = ' '.join(p for b in blocks if b['level'] == 2 and b['text'].startswith(('المبحث الرابع', 'المبحث الخامس')) for p in b['paras'])
js = set(sentences(jan_text)); ws_ = set(sentences(wadee_text))
report['identical_sentences_between_authors'] = sorted(js & ws_)[:10]

def shingles(t, k=5):
    w = norm(t).split()
    return {' '.join(w[i:i + k]) for i in range(max(0, len(w) - k + 1))}
sh_j, sh_w = shingles(jan_text), shingles(wadee_text)
inter = sh_j & sh_w
report['shared_5grams'] = len(inter)
report['shared_5grams_sample'] = sorted(inter)[:10]
report['overlap_ratio'] = round(len(inter) / max(1, min(len(sh_j), len(sh_w))), 4)

# ---------------------------------------------------------------- علامات لغوية
issues = []
pats = [(r'\bإسم\b', 'همزة قطع في «إسم»'),
        (r'\bهاذا\b', 'هاذا'),
        (r'\bالذى\b', 'ألف مقصورة'),
        (r'\bلاكن\b', 'لاكن'),
        (r'  +', 'مسافتان')]
for pat, desc in pats:
    m = re.findall(pat, all_text)
    if m:
        issues.append((desc, len(m)))
report['language_flags'] = issues

# ---------------------------------------------------------------- الهوامش: صيغة «المرجع نفسه»
report['footnotes_first_person_samples'] = footnotes[:3]
report['footnotes_containing_ibid'] = sum(1 for t in footnotes if 'المرجع نفسه' in t)

json.dump(report, open('build/report.json', 'w'), ensure_ascii=False, indent=2)
print(json.dumps(report, ensure_ascii=False, indent=2)[:4000])
