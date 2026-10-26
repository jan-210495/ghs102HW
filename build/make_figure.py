# -*- coding: utf-8 -*-
"""إنشاء الشكل (1): مسار انتقال الرأي من الحدث إلى الجمهور."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from matplotlib import font_manager as fm

FPATH = 'assets/fonts/amiri/Amiri-Regular.ttf'
fm.fontManager.addfont(FPATH)
FP = fm.FontProperties(fname=FPATH)
FB = fm.FontProperties(fname='assets/fonts/amiri/Amiri-Bold.ttf')

# ملاحظة: matplotlib يتولّى تشكيل الحروف العربية وترتيب الاتجاه بنفسه،
# ولذلك تُمرَّر النصوص كما هي دون معالجة مسبقة.
def ar(t):
    return t

fig, ax = plt.subplots(figsize=(7.4, 7.8), dpi=240)
ax.set_xlim(0, 10); ax.set_ylim(0, 13.3); ax.axis('off')

boxes = [
    ("١. الحدث أو القضية العامة", "واقعة تمسّ مصالح جماعة، وتحتاج إلى صياغة قبل أن تصل إلى الناس"),
    ("٢. البوابة: الوسيلة أو المنصة", "اختيار ما يُنشر، وترتيب الأولويات، وبناء إطار الفهم"),
    ("٣. الوسيط البشري", "قادة الرأي في الجماعات الصغيرة يعيدون الصياغة وينقلونها"),
    ("٤. مناخ الرأي", "قراءة موقف الأغلبية، والخوف من العزلة، ودوامة الصمت"),
    ("٥. التعبير المقيس", "استطلاع، أو تحليل مشاعر، أو ملاحظة ميدانية منظّمة"),
]

x0, w, h, gap = 1.75, 6.55, 1.95, 0.42
ys = []
for i, (title, sub) in enumerate(boxes):
    y = 13.0 - h - i * (h + gap)
    ys.append(y)
    ax.add_patch(FancyBboxPatch((x0, y), w, h,
                                boxstyle="round,pad=0.02,rounding_size=0.14",
                                linewidth=1.0, edgecolor='#333333', facecolor='#F3F3F3'))
    ax.text(x0 + w / 2, y + h - 0.62, ar(title), ha='center', va='center',
            fontsize=13, fontproperties=FB, color='#111111')
    ax.text(x0 + w / 2, y + 0.62, ar(sub), ha='center', va='center',
            fontsize=10.2, fontproperties=FP, color='#222222')

for i in range(len(boxes) - 1):
    ax.add_patch(FancyArrowPatch((x0 + w / 2, ys[i]), (x0 + w / 2, ys[i + 1] + h + 0.03),
                                 arrowstyle='-|>', mutation_scale=15, linewidth=1.1,
                                 color='#444444'))

ax.add_patch(FancyArrowPatch((x0 - 0.08, ys[4] + h / 2), (x0 - 0.08, ys[0] + h / 2),
                             connectionstyle="arc3,rad=0.16", arrowstyle='-|>',
                             mutation_scale=15, linewidth=1.0, color='#808080',
                             linestyle=(0, (4, 3))))
ax.text(5.0, 0.42, ar("الخط المتقطّع: تغذية راجعة، إذ تعيد نتائج القياس تشكيل القضية نفسها"),
        ha='center', va='center', fontsize=9.5, fontproperties=FP, color='#555555')

plt.subplots_adjust(left=0.02, right=0.98, top=0.99, bottom=0.01)
plt.savefig('assets/fig1.png', bbox_inches='tight', facecolor='white')
print('figure saved')
