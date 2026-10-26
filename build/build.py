# -*- coding: utf-8 -*-
"""بناء ملف الوظيفة النهائي (.docx)."""
import re
import sys
import json
import pathlib

sys.path.insert(0, 'build')
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxgen import ArabicDocx, para_rtl, run_rtl, table_rtl, clean, FONT, FONT_SIZE
import data

ROOT = pathlib.Path('.')
OUT = ROOT / 'الوظيفة_الأولى_الرأي_العام_جان_وديع.docx'

HEADER_TEXT = 'مقرر المدخل إلى علم الاجتماع — الوظيفة الأولى: الرأي العام — العام الدراسي 2026/2027'

d = ArabicDocx()
doc = d.doc

# إعدادات عامة: لغة عربية افتراضية
_settings = doc.settings.element
theme_fonts = None

# ------------------------------------------------------------------ الغلاف
def cover_line(text, size=14, bold=False, space_after=6, space_before=0, color=None):
    text = clean(text)
    p = doc.add_paragraph()
    para_rtl(p, align='center', indent=0, line=1.15, before=space_before, after=space_after)
    run_rtl(p.add_run(text), size=size, bold=bold, color=color)
    return p

cover_line('الجامعة الافتراضية السورية', size=20, bold=True, space_after=2)
cover_line('(SVU — الجامعة الافتراضية السورية)', size=11, space_after=14, color=RGBColor(0x55, 0x55, 0x55))
cover_line('كلية الهندسة المعلوماتية — برنامج الإجازة في تقانة المعلومات', size=14, space_after=4)
cover_line('(بكالوريوس تقانات المعلومات — BAIT، ويُشار إليه في هذا الغلاف اختصاراً بالحروف اللاتينية)', size=11, space_after=6, color=RGBColor(0x55, 0x55, 0x55))
cover_line('مقرر: المدخل إلى علم الاجتماع', size=16, bold=True, space_before=18, space_after=2)
cover_line('الوظيفة الأولى', size=20, bold=True, space_before=10, space_after=2)
cover_line('«الرأي العام»', size=22, bold=True, space_after=16)
cover_line('إعداد: أساتذة مادة علم الاجتماع', size=13, space_after=2)
cover_line('تاريخ التوزيع: 14/10/2026 — تاريخ الإعادة: 12/11/2026', size=13, space_after=18)

# جدول بيانات الطالبين على الغلاف
t = doc.add_table(rows=3, cols=2)
t.style = 'Table Grid'
t.alignment = WD_TABLE_ALIGNMENT.CENTER
table_rtl(t)
cov = [
    ['اسم الطالب', 'الرقم الجامعي والمسار الدراسي'],
    ['جان شادي عبود', 'jan_436439 — المعهد التقاني للحاسوب، ثم الإجازة في تقانة المعلومات'],
    ['وديع هيثم سلوم', 'wadee_440350 — المعهد التقاني للحاسوب، ثم الإجازة في تقانة المعلومات'],
]
for ri, row in enumerate(cov):
    for ci, val in enumerate(row):
        cell = t.cell(ri, ci)
        cell.text = ''
        p = cell.paragraphs[0]
        para_rtl(p, align='center', indent=0, line=1.1, before=3, after=3)
        run_rtl(p.add_run(clean(val)), size=12, bold=(ri == 0))
        cell.width = Cm(8.2)

cover_line('العام الدراسي 2026/2027', size=14, bold=True, space_before=22, space_after=2)
cover_line('دمشق — الجمهورية العربية السورية', size=12, space_after=0, color=RGBColor(0x55, 0x55, 0x55))

# ------------------------------------------------------------------ مقطع جديد للجسم
d.new_section(restart_numbering=True)
d.add_header_footer(HEADER_TEXT)

# ------------------------------------------------------------------ إقرار الأصالة
d.heading('إقرار بالأصالة', level=1)
d.para('نحن الموقّعان أدناه، جان شادي عبود (الرقم الجامعي: jan_436439)، ووديع هيثم سلوم (الرقم الجامعي: wadee_440350)، نُقرّ بما يلي:', indent=1.0)
for item in [
    'أنّ هذه الوظيفة بعنوان «الرأي العام» هي عمل أصيل أعددناه بأنفسنا لمقرر «المدخل إلى علم الاجتماع»، ولم يُقدَّم من قبل في أي مقرر أو جامعة أخرى، ولم يُنشر بأي صيغة.',
    'أنّ كل ما نقلناه من كتب أو مقالات أو تقارير أو مواقع إلكترونية موثّق في موضعه بهامش سفلي، ومُدرَج في مسرد المراجع والمصادر في نهاية البحث.',
    'أنّ ما ورد فيها من أمثلة من بيئة العمل والمجموعة الكشفية هو من خبرتنا الشخصية، وقد وُصف بصدق على أنه ملاحظة محدودة غير ممثّلة إحصائياً، ولم نخترع أي نسبة أو عيّنة أو نتيجة استبيان.',
    'أنّنا نعلم أنّ أي تشابه بين هذه الوظيفة وغيرها من الوظائف يُعرّضنا لعقوبة الصفر في الوظيفة والحرمان من التقديم للامتحان، ونقبل هذه العقوبة إن ثبت التشابه.',
    'أنّ العمل جرى ضمن فريق من طالبين، وأنّ توزيع المهام موضّح في الجدول (1) وفي المقدمة، وأنّ كل باحث يتحمّل مسؤولية القسم الذي كتبه.',
]:
    d.para('• ' + item, indent=0.6, after=4)

d.para('', after=10)
sig = doc.add_table(rows=2, cols=2)
sig.style = 'Table Grid'
table_rtl(sig)
sig_rows = [['الباحث الأول', 'الباحث الثاني'],
            ['جان شادي عبود\njan_436439\nالتوقيع: ....................\nالتاريخ: 12/11/2026',
             'وديع هيثم سلوم\nwadee_440350\nالتوقيع: ....................\nالتاريخ: 12/11/2026']]
for ri, row in enumerate(sig_rows):
    for ci, val in enumerate(row):
        cell = sig.cell(ri, ci)
        cell.text = ''
        parts = val.split('\n')
        p = cell.paragraphs[0]
        para_rtl(p, align='center', indent=0, line=1.2, before=4, after=4)
        run_rtl(p.add_run(clean(parts[0])), size=12, bold=(ri == 0))
        for extra in parts[1:]:
            p2 = cell.add_paragraph()
            para_rtl(p2, align='center', indent=0, line=1.2, before=0, after=4)
            run_rtl(p2.add_run(clean(extra)), size=12)

# ------------------------------------------------------------------ شكر وتقدير
d.heading('شكر وتقدير', level=1, page_break=True)
d.para('نشكر أساتذة مادة «المدخل إلى علم الاجتماع» في الجامعة الافتراضية السورية على ما قدّموه من مادة علمية ومن توجيهات منهجية أسهمت في ترتيب أفكارنا، ونشكر كل من ساعدنا في الوصول إلى مصدر أو مراجعة صياغة. ونخصّ بالذكر زملاءنا في العمل وفي المجموعتين الكشفيّتين، الذين كان حديثهم اليومي مادة خاماً لكثير من الأمثلة الواردة في هذه الوظيفة. وكل ما ورد فيها من خطأ فهو منّا، وما أصاب ففضلٌ لمن علّمنا.')

# ------------------------------------------------------------------ فهرس المحتويات
p = doc.add_paragraph()
para_rtl(p, align='center', indent=0, line=1.3, before=0, after=10)
run_rtl(p.add_run('فهرس المحتويات'), size=18, bold=True)
d.para('(فهرس آلي قابل للتحديث: يُحدَّث تلقائياً عند فتح الملف في برنامج وورد، أو يدوياً بتحديد الحقل ثم اختيار «تحديث الحقل» من قائمة المراجع؛ وأرقام الصفحات في فهرس الجداول مرتبطة بحقول إحالة تتحدّث مع الفهرس نفسه)', indent=0, size=11)
d.toc_field('1-2')

# ------------------------------------------------------------------ فهرس الجداول والأشكال
d.heading('فهرس الجداول والأشكال', level=1, page_break=True)
d.para('يُدرج هذا الفهرس يدوياً وتُحدَّث أرقام صفحاته آلياً بحقول الإحالة، وهو يشير إلى مواضع الجداول والأشكال داخل البحث:', indent=0)

LIST_ITEMS = [
    ('الجدول (1): توزيع العمل بين الباحثين في الوظيفة الأولى', '_Tbl1'),
    ('الجدول (2): مقارنة بين الاتجاهات الثلاثة في تعريف الرأي العام', '_Tbl2'),
    ('الجدول (3): التمييز بين الرأي العام والمفاهيم المجاورة', '_Tbl3'),
    ('الجدول (4): خصائص الرأي العام ومؤشرات قياسها', '_Tbl4'),
    ('الجدول (5): مقارنة النظريات المفسّرة لتشكّل الرأي العام', '_Tbl5'),
    ('الشكل (1): مسار انتقال الرأي من الحدث إلى الجمهور', '_Fig1'),
    ('الجدول (6): أدوات قياس الرأي العام في المجال الاقتصادي وحدودها', '_Tbl7'),
    ('الجدول (7): مسرد المصطلحات الأساسية الواردة في الوظيفة', '_Tbl8'),
]
for label, bm in LIST_ITEMS:
    p = doc.add_paragraph()
    para_rtl(p, align='both', indent=0, line=1.4, after=8)
    # نقاط القيادة
    pPr = p._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    for pos, val in ((9000, 'right'),):
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), 'right')
        tab.set(qn('w:leader'), 'dot')
        tab.set(qn('w:pos'), str(pos))
        tabs.append(tab)
    pPr.insert_element_before(tabs, 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
                              'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE',
                              'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid',
                              'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
                              'w:suppressOverlap', 'w:jc', 'w:textDirection',
                              'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl',
                              'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange')
    run_rtl(p.add_run(clean(label) + '\t'), size=13)
    d.pageref_field(p, bm, size=13)

# ------------------------------------------------------------------ جدول توزيع العمل
d.heading('جدول توزيع العمل بين الباحثين', level=1, page_break=True)
d.table(data.T1['rows'], widths=data.T1['widths'], font_size=11,
        caption=data.T1['caption'], bookmark=data.T1['bookmark'])
d.para('يوضّح الجدول أعلاه أنّ العمل جرى في ثلاث طبقات: مهام فردية يتحمّل صاحبها مسؤولية صياغتها وأمثلتها، ومهام مشتركة تُراجع من الطرفين. وقد حرصنا على ألا يقلّ نصيب أي منّا عن خمس صفحات، وعلى أن يكون لكل قسم صوت مختلف في المفردات وبناء الجملة، مع الحفاظ على وحدة النسق الشكلي والمنهجي.', indent=1.0)

# ------------------------------------------------------------------ المتن
FILES = ['content/01_intro.md', 'content/02_q1.md', 'content/03_q2.md',
         'content/04_q3.md', 'content/05_q4.md', 'content/06_q5.md',
         'content/07_back.md']

TABLE_TOKENS = {'@T1': data.T1, '@T2': data.T2, '@T3': data.T3, '@T4': data.T4,
                '@T5': data.T5, '@T7': data.T7, '@T8': data.GLOSSARY}

first_h1 = {'done': False}


def emit_block(block):
    block = block.strip('\n')
    if not block.strip():
        return
    m = re.match(r'^@(T\d|F\d)$', block.strip())
    if m:
        tok = '@' + m.group(1)
        if tok == '@F1':
            d.figure('assets/fig1.png', width_cm=13.2,
                     caption='الشكل (1): مسار انتقال الرأي من الحدث إلى الجمهور (من إعداد الباحثين)',
                     bookmark='_Fig1')
        else:
            tb = TABLE_TOKENS[tok]
            d.table(tb['rows'], widths=tb['widths'], font_size=10.5,
                    caption=tb['caption'], bookmark=tb['bookmark'])
        return
    if block.startswith('@BY '):
        p = doc.add_paragraph()
        para_rtl(p, align='right', indent=0, line=1.15, before=0, after=10)
        run_rtl(p.add_run(clean(block[4:].strip())), size=12, bold=True)
        return
    if block.startswith('### '):
        d.heading(block[4:].strip(), level=2)
        return
    if block.startswith('## '):
        d.heading(block[3:].strip(), level=2)
        return
    if block.startswith('# '):
        d.heading(block[2:].strip(), level=1, page_break=True)
        first_h1['done'] = True
        return
    d.para(block.strip())


for f in FILES:
    txt = (ROOT / f).read_text(encoding='utf-8')
    for block in re.split(r'\n\s*\n', txt):
        emit_block(block)

# ------------------------------------------------------------------ المسارد
d.heading('مسرد المصطلحات', level=1, page_break=True)
d.para('جمعنا في هذا المسرد المصطلحات التي تكرّر استخدامها في البحث، وحرصنا على أن يكون تعريف كل مصطلح تعريفاً إجرائياً يصلح للاستعمال في القياس، لا تعريفاً معجمياً عاماً:', indent=1.0)
d.table(data.GLOSSARY['rows'], widths=data.GLOSSARY['widths'], font_size=11,
        caption=data.GLOSSARY['caption'], bookmark=data.GLOSSARY['bookmark'])

# ------------------------------------------------------------------ المراجع
d.heading('مسرد المراجع والمصادر', level=1, page_break=True)
d.para('رتّبنا المراجع هجائياً داخل كل قائمة من القوائم الأربع الآتية، وأثبتنا لكل مرجع بياناته الببليوغرافية كاملة، وأضفنا المعرّف الرقمي أو الرابط مع تاريخ الاسترجاع حيث يكون المرجع إلكترونياً. وقد تحقّقنا من بيانات المؤلف والعنوان والسنة والدورية لكل مرجع قبل إدراجه، ولم نُدرج أي مرجع لم نتأكّد من وجوده.', indent=1.0)

def ref_list(title, items, note=None):
    d.heading(title, level=2)
    if note:
        d.para(note, indent=1.0, after=6)
    for i, (author, title_, rest) in enumerate(items, start=1):
        p = doc.add_paragraph()
        para_rtl(p, align='both', indent=0, line=1.3, before=0, after=8)
        p.paragraph_format.left_indent = Cm(0.8)
        run_rtl(p.add_run(f'{i}. '), size=13)
        run_rtl(p.add_run(clean(author) + ' '), size=13, bold=True)
        run_rtl(p.add_run(clean('«' + title_ + '»، ')), size=13, italic=True)
        run_rtl(p.add_run(clean(rest)), size=13)

ref_list('أولاً: المراجع العربية', data.REFS_ARABIC)
ref_list('ثانياً: المراجع الأجنبية (مرتّبة هجائياً بالعنوان العربي، ويُذكر العنوان الأصلي بين قوسين)',
         data.REFS_FOREIGN)
ref_list('ثالثاً: التقارير المؤسسية', data.REFS_INSTITUTIONAL)
ref_list('رابعاً: المصادر الإلكترونية',
         [('الجامعة الافتراضية السورية، وحدة «الرأي العام وطرق قياسه».',
           'الوحدة التعليمية الثانية والثالثة (ملف إلكتروني)',
           'دمشق: الجامعة الافتراضية السورية. الرابط: '
           'https://pedia.svuonline.org/pluginfile.php/2858/mod_resource/content/30/ — '
           'تاريخ الاسترجاع: 2026/10/26.'),
          ('مجلة الجزيرة لدراسات الاتصال والإعلام.', 'العدد الثالث (يناير 2024)',
           'مركز الجزيرة للدراسات، الدوحة. الرابط: https://aljazeerajournal.aljazeera.net/ — '
           'تاريخ الاسترجاع: 2026/10/26.')])

d.para('ملاحظة توثيقية: اعتمدنا في المتن صيغة (اسم المؤلف، السنة: الصفحة) في الإشارة المختصرة داخل الجدول (2) وما يليه عند الحاجة، أما التوثيق الكامل فقد وُضع في الهوامش السفلية عند أول إشارة إلى كل مرجع، ثم بصيغة «المرجع نفسه» في الإشارات التالية، حفاظاً على شرط الوظيفة.', indent=1.0)

# ------------------------------------------------------------------ الملاحق
d.heading('الملاحق', level=1, page_break=True)
d.heading('الملحق (أ): أداة استبيان مقترحة (لم تُنفَّذ)', level=2)
d.para(data.APPENDIX_A_INTRO, indent=1.0)
for i, q in enumerate(data.SURVEY, start=1):
    p = doc.add_paragraph()
    para_rtl(p, align='both', indent=0, line=1.3, after=6)
    p.paragraph_format.left_indent = Cm(0.7)
    run_rtl(p.add_run(f'{i}. '), size=13)
    run_rtl(p.add_run(clean(q)), size=13)

d.heading('الملحق (ب): محاور المقابلة شبه الموجّهة', level=2)
d.para(data.APPENDIX_B_NOTE, indent=1.0)
for ax, q in data.INTERVIEW_AXES:
    d.para('**' + ax + ':** ' + q, indent=1.0, after=6)

d.heading('الملحق (ج): مراجع ومصادر تحتاج تحقّق الطالب قبل التسليم', level=2)
d.para(data.APPENDIX_C_INTRO, indent=1.0)
for it in data.APPENDIX_C_ITEMS:
    d.para('• ' + it, indent=0.6, after=5)

# ------------------------------------------------------------------ الحفظ
for section in doc.sections[1:]:
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)

doc.core_properties.title = 'الوظيفة الأولى: الرأي العام — مقرر المدخل إلى علم الاجتماع'
doc.core_properties.author = 'جان شادي عبود؛ وديع هيثم سلوم'
doc.core_properties.subject = 'الجامعة الافتراضية السورية — الإجازة في تقانة المعلومات'
doc.core_properties.language = 'ar'
d.save(str(OUT))
print('saved', OUT)
