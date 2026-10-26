import sys; sys.path.insert(0,'build')
from docxgen import ArabicDocx
d = ArabicDocx()
d.add_header_footer("مقرر المدخل إلى علم الاجتماع — الوظيفة الأولى: الرأي العام")
d.heading("الفصل الأول: مدخل تجريبي")
d.para("هذا نصٌّ تجريبي للتحقق من عمل الهوامش{{والتر ليبمان، الرأي العام (Public Opinion)، نيويورك: هاركوت بريس، 1922، ص 3.}} ثم نص تالٍ مع كلمة _مائلة_ وأخرى *عريضة*.")
d.toc_field()
d.table([["الخاصية","المعنى","مؤشر"],["العمومية","تعمّ القضية جمهوراً واسعاً","حجم الجمهور{{المرجع نفسه، ص 12.}}"]], widths=[4,6,5], caption="الجدول (1): تجربة", bookmark="_Tbl1")
p = d.doc.add_paragraph(); from docxgen import para_rtl
para_rtl(p, align='center'); from docxgen import run_rtl
r=p.add_run(); run_rtl(r); d._add_field(r, "PAGEREF _Tbl1 \\h", "—")
d.heading("الثاني", level=2)
d.para("فقرة ثالثة.")
d.save('/tmp/test.docx')
print("saved")
