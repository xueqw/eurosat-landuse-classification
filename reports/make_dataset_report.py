"""Builds Dataset_Selection_Report.docx from the executed notebook's figures and numbers."""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FIG = "figures/"
doc = Document()

# ---------- parallel PDF output (reportlab) ----------
from xml.sax.saxutils import escape
from PIL import Image as PILImage
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image as RLImage,
                                Table as RLTable, TableStyle, KeepTogether)

FD = "/System/Library/Fonts/Supplemental/"
pdfmetrics.registerFont(TTFont("Arial", FD + "Arial.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", FD + "Arial Bold.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Italic", FD + "Arial Italic.ttf"))
pdfmetrics.registerFont(TTFont("Arial-BoldItalic", FD + "Arial Bold Italic.ttf"))
pdfmetrics.registerFont(TTFont("CourierNew", FD + "Courier New.ttf"))
from reportlab.pdfbase.pdfmetrics import registerFontFamily
registerFontFamily("Arial", normal="Arial", bold="Arial-Bold", italic="Arial-Italic", boldItalic="Arial-BoldItalic")

NAVY = colors.HexColor("#1F3A5F")
S_BODY = ParagraphStyle("body", fontName="Arial", fontSize=10.5, leading=14, spaceAfter=6)
S_H1 = ParagraphStyle("h1", fontName="Arial-Bold", fontSize=13.5, leading=17, textColor=NAVY, spaceBefore=12, spaceAfter=5)
S_H2 = ParagraphStyle("h2", fontName="Arial-Bold", fontSize=11.5, leading=15, textColor=NAVY, spaceBefore=7, spaceAfter=3)
S_BUL = ParagraphStyle("bul", parent=S_BODY, leftIndent=16, bulletIndent=5, spaceAfter=3)
S_CAP = ParagraphStyle("cap", fontName="Arial-Italic", fontSize=9.5, leading=12, alignment=TA_CENTER, spaceAfter=6)
S_CELL = ParagraphStyle("cell", fontName="Arial", fontSize=9, leading=11.5)
S_CODE = ParagraphStyle("code", fontName="CourierNew", fontSize=8.2, leading=10)
story = []
TEXT_W = 6.5 * inch


def rl_markup(text, bold_lead=None, italic=False):
    s = (f"<b>{escape(bold_lead)}</b>" if bold_lead else "") + escape(text)
    return f"<i>{s}</i>" if italic else s

# ---------- base styles ----------
st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(11)
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.15
for lvl, size in ((1, 14), (2, 12)):
    h = doc.styles[f"Heading {lvl}"]
    h.font.name = "Calibri"
    h.font.size = Pt(size)
    h.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    h.paragraph_format.space_before = Pt(12 if lvl == 1 else 8)
    h.paragraph_format.space_after = Pt(4)
for s in doc.sections:
    s.left_margin = s.right_margin = Inches(1)
    s.top_margin = s.bottom_margin = Inches(0.9)


def heading(text, level):
    doc.add_heading(text, level=level)
    story.append(Paragraph(escape(text), S_H1 if level == 1 else S_H2))


def para(text="", bold_lead=None, italic=False, align=None, size=None, pdf=True, bold=False):
    if pdf:
        st_ = ParagraphStyle("p", parent=S_BODY, fontSize=(size or 11) - 0.5,
                             leading=((size or 11) - 0.5) * 1.33,
                             alignment=TA_CENTER if align == WD_ALIGN_PARAGRAPH.CENTER else 0)
        m = rl_markup(text, bold_lead, italic)
        story.append(Paragraph(f"<b>{m}</b>" if bold else m, st_))
    p = doc.add_paragraph()
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
    r = p.add_run(text)
    r.italic = italic
    if bold:
        for run in p.runs:
            run.bold = True
    if size:
        for run in p.runs:
            run.font.size = Pt(size)
    if align:
        p.alignment = align
    return p


def bullet(text, bold_lead=None):
    story.append(Paragraph(rl_markup(text, bold_lead), S_BUL, bulletText="•"))
    p = doc.add_paragraph(style="List Bullet")
    if bold_lead:
        p.add_run(bold_lead).bold = True
    p.add_run(text)
    p.paragraph_format.space_after = Pt(2)
    return p


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def table(rows, widths=None, header=True, bold_last=False, font=9.5):
    cs = ParagraphStyle("c", parent=S_CELL, fontSize=font - 0.5, leading=(font - 0.5) * 1.28)
    data = []
    for i, row in enumerate(rows):
        b = (header and i == 0) or (bold_last and i == len(rows) - 1)
        data.append([Paragraph(f"<b>{escape(str(v))}</b>" if b else escape(str(v)), cs) for v in row])
    scale = TEXT_W / (sum(widths) * inch) if widths else 1
    t_ = RLTable(data, colWidths=[w * inch * min(scale, 1) for w in widths] if widths else None,
                 hAlign="CENTER", repeatRows=1 if header else 0)
    ts = [("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#9AA5B1")),
          ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]
    if header:
        ts.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF5")))
    t_.setStyle(TableStyle(ts))
    if story and isinstance(story[-1], Paragraph) and story[-1].style.name == "cap":
        story.append(KeepTogether([story.pop(), t_]))   # caption stays with its table
    else:
        story.append(t_)
    story.append(Spacer(1, 8))
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i, j)
            c.text = ""
            r = c.paragraphs[0].add_run(str(val))
            r.font.size = Pt(font)
            c.paragraphs[0].paragraph_format.space_after = Pt(0)
            if (header and i == 0) or (bold_last and i == len(rows) - 1):
                r.bold = True
            if header and i == 0:
                shade(c, "E8EEF5")
            if widths:
                c.width = Inches(widths[j])
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


def caption(text, pdf=True):
    if pdf:
        story.append(Paragraph(escape(text), S_CAP))
    p = para(text, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=9.5, pdf=False)
    return p


def figure(path, width, cap):
    w, h = PILImage.open(path).size
    img = RLImage(path, width=width * inch, height=width * inch * h / w)
    story.append(KeepTogether([img, Spacer(1, 3), Paragraph(escape(cap), S_CAP)]))
    doc.add_picture(path, width=Inches(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption(cap, pdf=False)


def code_block(text, size=8.5):
    lines = "<br/>".join(escape(l).replace(" ", "&nbsp;") for l in text.split("\n"))
    cb = RLTable([[Paragraph(lines, S_CODE)]], colWidths=[TEXT_W])
    cb.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F5F7")),
                            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#C5CBD3")),
                            ("LEFTPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    if story and isinstance(story[-1], Paragraph):   # keep the lead-in line with its code box
        story.append(KeepTogether([story.pop(), cb]))
    else:
        story.append(cb)
    story.append(Spacer(1, 8))
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.cell(0, 0)
    shade(c, "F4F5F7")
    c.text = ""
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    r.font.name = "Menlo"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Menlo")
    r.font.size = Pt(size)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


# ---------- title block ----------
p = para("CPS 4841 Computer Vision — Course Project", align=WD_ALIGN_PARAGRAPH.CENTER, size=16, bold=True)
p = para("Dataset Selection Report: Land-Use Classification of Satellite Images (EuroSAT)",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=12.5)
para("Student: Qiwen Xue    ·    Section 03    ·    November 16, 2026",
     align=WD_ALIGN_PARAGRAPH.CENTER, size=10.5)

# ---------- 1. Introduction ----------
heading("1. Introduction to the Chosen Problem", level=1)
heading("a. Problem description", level=2)
para("The task is land-use and land-cover (LULC) classification of satellite images. Each input is a "
     "64 × 64 pixel RGB patch cut from a Sentinel-2 satellite image (10 m per pixel, so one patch covers "
     "about 640 m × 640 m of ground). The model must assign exactly one of ten land-use classes to the "
     "whole patch: AnnualCrop, Forest, HerbaceousVegetation, Highway, Industrial, Pasture, PermanentCrop, "
     "Residential, River or SeaLake. This is a single-label, 10-class image classification problem "
     "(Task 1 in the project appendix, “classifying land use from satellite images”).")
heading("b. Why I chose this topic", level=2)
para("Satellite imagery is one of the largest and fastest-growing sources of visual data, and most of it is "
     "never looked at by a person. I wanted a problem where computer vision does real work at a scale "
     "humans cannot. EuroSAT is also a good fit for a five-week project: it is a clean, well-documented, "
     "fully labeled benchmark of a manageable size (27,000 small images, about 90 MB), it runs comfortably "
     "on the free Colab GPU, and it allows a fair comparison between a traditional pipeline (hand-crafted "
     "colour and texture features) and deep learning. Several classes look similar from above "
     "(AnnualCrop vs. PermanentCrop, Pasture vs. HerbaceousVegetation, Highway vs. River), so the "
     "comparison should be informative rather than trivial.")
heading("c. Why it is worthy of research", level=2)
bullet("Land-cover maps are used for agricultural monitoring, urban planning, deforestation tracking, "
       "flood and disaster response, and climate reporting. Automatic classification makes it possible to "
       "update these maps frequently and over whole countries.", "Practical impact. ")
bullet("Sentinel-2 images are released free of charge by the European Space Agency (Copernicus programme) "
       "every few days, so a working classifier can be applied to new data immediately.", "Open data. ")
bullet("Aerial images differ from everyday photos: there is no fixed “up” direction, objects are tiny, and "
       "texture and colour carry most of the information. This makes it a good test of how well models "
       "pre-trained on ImageNet transfer to a different domain, and of whether simple colour and texture "
       "features are enough for some classes.", "Open research questions. ")

# ---------- 2. Dataset overview ----------
heading("2. Dataset Overview and Visualization", level=1)
heading("a. Overview", level=2)
table([
    ["Property", "Value"],
    ["Name", "EuroSAT (RGB version)"],
    ["Source / paper", "Helber, Bischke, Dengel & Borth, IEEE JSTARS 12(7), 2019 [1]"],
    ["Download link", "https://zenodo.org/records/7711810  (file EuroSAT_RGB.zip, 94.7 MB)"],
    ["Mirrors", "https://github.com/phelber/EuroSAT ; also built into torchvision (datasets.EuroSAT)"],
    ["License", "MIT (redistribution allowed; the repository will still use a download script, not the data)"],
    ["Imagery", "Sentinel-2 satellite, 10 m ground resolution, 34 European countries"],
    ["Number of images", "27,000"],
    ["Image format", "JPEG, 64 × 64 pixels, 3 channels (RGB), 8-bit"],
    ["Number of classes", "10 (single label per image)"],
    ["Annotation format", "Class name = folder name (e.g. Forest/Forest_12.jpg)"],
    ["Official train/test split", "None — a stratified split is defined in Section 2c"],
], widths=[1.7, 4.8])

para("Table 2 lists every class with its number of images, measured directly from the downloaded files "
     "(notebook cell 2).")
caption("Table 2. Classes and number of images per class")
table([
    ["ID", "Class", "What the patch shows", "Images", "%"],
    ["0", "AnnualCrop", "Fields of crops replanted every year (cereals, maize, …)", "3,000", "11.11"],
    ["1", "Forest", "Continuous tree cover", "3,000", "11.11"],
    ["2", "HerbaceousVegetation", "Natural grassland, shrubs, low vegetation", "3,000", "11.11"],
    ["3", "Highway", "Motorways and major roads crossing the patch", "2,500", "9.26"],
    ["4", "Industrial", "Factories, warehouses, large flat roofs", "2,500", "9.26"],
    ["5", "Pasture", "Grassland used for grazing livestock", "2,000", "7.41"],
    ["6", "PermanentCrop", "Orchards, vineyards, olive groves", "2,500", "9.26"],
    ["7", "Residential", "Housing areas with dense small buildings", "3,000", "11.11"],
    ["8", "River", "Rivers and their banks", "2,500", "9.26"],
    ["9", "SeaLake", "Open water (sea or lakes)", "3,000", "11.11"],
    ["", "Total", "", "27,000", "100.00"],
], widths=[0.35, 1.6, 3.35, 0.65, 0.55], bold_last=True)

heading("b. Class distribution", level=2)
figure(FIG + "class_distribution.png", 6.5,
       "Figure 1. Number of images per class (left) and share of the dataset (right).")
para("The dataset is close to balanced. The largest classes have 3,000 images and the smallest (Pasture) has "
     "2,000, an imbalance ratio of only 1.5 : 1. Every class has at least 2,000 examples, which is enough to "
     "train a small CNN from scratch. Because the classes are not perfectly equal, I will report macro-averaged "
     "metrics in addition to overall accuracy, so that the smaller classes count as much as the larger ones.")

figure(FIG + "mean_rgb.png", 6.0,
       "Figure 2. Mean R, G and B pixel value per class (300 random images per class).")
para("Figure 2 shows that simple colour statistics already separate some classes: Forest and SeaLake are "
     "dark with a blue-green tint, while AnnualCrop, PermanentCrop and Industrial are much brighter. Other "
     "pairs, such as Highway and Residential or Pasture and River, have almost the same average colour and "
     "will need texture or shape information. This is useful for planning the methods: a colour-histogram "
     "baseline should do well on some classes and poorly on others.")

heading("c. Planned train / validation / test split", level=2)
para("EuroSAT does not ship with an official split, so I created a stratified 70 % / 15 % / 15 % split "
     "with a fixed random seed (42). Every class keeps the same proportion in each part. The split is "
     "saved to splits.csv so that all three methods are trained and tested on exactly the same images.")
caption("Table 3. Number of images per class in each split")
table([
    ["Class", "Train", "Validation", "Test", "Total"],
    ["AnnualCrop", "2,100", "450", "450", "3,000"],
    ["Forest", "2,100", "450", "450", "3,000"],
    ["HerbaceousVegetation", "2,100", "450", "450", "3,000"],
    ["Highway", "1,750", "375", "375", "2,500"],
    ["Industrial", "1,750", "375", "375", "2,500"],
    ["Pasture", "1,400", "300", "300", "2,000"],
    ["PermanentCrop", "1,750", "375", "375", "2,500"],
    ["Residential", "2,100", "450", "450", "3,000"],
    ["River", "1,750", "375", "375", "2,500"],
    ["SeaLake", "2,100", "450", "450", "3,000"],
    ["Total", "18,900", "4,050", "4,050", "27,000"],
], widths=[2.0, 1.0, 1.0, 1.0, 1.0], bold_last=True)

# ---------- 3. Proof that it works ----------
heading("3. Proof That the Dataset Works", level=1)
para("I wrote a notebook, 01_dataset_check.ipynb, that downloads the zip file from Zenodo, unzips it, opens "
     "every image, counts the classes, makes the plots in this report and displays sample images with their "
     "labels. It uses only libraries that come preinstalled in Google Colab (NumPy, pandas, Pillow, "
     "Matplotlib). The core loading code and its printed output are shown below.")
code_block(
    "URL  = 'https://zenodo.org/records/7711810/files/EuroSAT_RGB.zip'\n"
    "urllib.request.urlretrieve(URL, 'data/EuroSAT_RGB.zip')\n"
    "zipfile.ZipFile('data/EuroSAT_RGB.zip').extractall('data')\n"
    "ROOT = 'data/EuroSAT_RGB'\n"
    "CLASSES = sorted(os.listdir(ROOT))\n"
    "for label_id, cls in enumerate(CLASSES):\n"
    "    for path in glob.glob(f'{ROOT}/{cls}/*.jpg'):\n"
    "        with Image.open(path) as im:          # opens every file\n"
    "            shapes[(im.size, im.mode)] += 1\n"
    "        records.append({'path': path, 'label': cls, 'label_id': label_id})")
caption("Code output (notebook cells 1–2):")
code_block(
    "Dataset folder: data/EuroSAT_RGB\n"
    "Class folders: ['AnnualCrop', 'Forest', 'HerbaceousVegetation', 'Highway', 'Industrial',\n"
    "                'Pasture', 'PermanentCrop', 'Residential', 'River', 'SeaLake']\n"
    "Number of classes: 10\n"
    "Total images     : 27000\n"
    "Unreadable files : 0\n"
    "Image size / mode: {((64, 64), 'RGB'): 27000}")
figure(FIG + "sample_grid.png", 6.5,
       "Figure 3. Notebook output: three random images from each class, with the class name and label id "
       "above each column and the file name below each image.")
para("All 27,000 files were downloaded and opened without errors, every image has the expected size "
     "(64 × 64, RGB), and the samples in Figure 3 visibly match their labels: fields in AnnualCrop, dense "
     "trees in Forest, large flat roofs in Industrial, water in River and SeaLake. The link works and the "
     "format is usable, so there is no need to change datasets.")

# ---------- 4. Annotation verification ----------
heading("4. Annotation Verification", level=1)
heading("a. Verifying that the annotations exist and are consistent", level=2)
para("Notebook cell 6 checks the annotations automatically. Its output:")
code_block(
    "Images with a label        : 27000 of 27000\n"
    "Folder/file-name mismatches: 0\n"
    "Duplicate file paths       : 0\n"
    "Byte-identical duplicates  : 0\n"
    "Label ids                  : {0: 'AnnualCrop', 1: 'Forest',\n"
    "                              2: 'HerbaceousVegetation', 3: 'Highway',\n"
    "                              4: 'Industrial', 5: 'Pasture', 6: 'PermanentCrop',\n"
    "                              7: 'Residential', 8: 'River', 9: 'SeaLake'}")
bullet("Every one of the 27,000 images has a class label.", "Complete: ")
bullet("Each label is stored twice, as the folder name and as the file-name prefix (Forest/Forest_12.jpg). "
       "The two agree for all images.", "Consistent: ")
bullet("No image appears twice, so the same picture cannot end up in both the training and the test set.",
       "No duplicates: ")
bullet("I looked through the sample grid (Figure 3) and further random samples; the labels match what the "
       "images show. According to the authors [1], the labels were assigned from existing land-cover "
       "maps and then checked manually.", "Visually correct: ")

heading("b. Annotation type and how it supports the task", level=2)
para("EuroSAT provides one image-level class label per image (10 mutually exclusive classes). It has no "
     "bounding boxes or segmentation masks, and none are needed: the task is to classify the whole patch, "
     "so the class label is exactly the ground truth required. The labels support the project as follows:")
bullet("Methods are trained with the labels (cross-entropy loss for the CNNs, class targets for the SVM).",
       "Training: ")
bullet("Predicted classes on the held-out test set are compared with the true labels to compute "
       "accuracy, per-class and macro-averaged precision, recall and F1-score, and a confusion matrix. "
       "The confusion matrix will show which similar-looking classes are confused.", "Evaluation: ")
bullet("Because the labels are per class, I can analyse results class by class and select misclassified "
       "examples for the error analysis required in the final report.", "Error analysis: ")
para("Known limitation: each patch carries one label even when it contains more than one land use "
     "(for example a highway passing through farmland). Some errors on such mixed patches may be label "
     "ambiguity rather than model failure; I will keep this in mind during the error analysis.")

# ---------- 5. Planned methods ----------
heading("5. Planned Methods and Metrics (Preliminary)", level=1)
para("This plan may change after the midterm, but it shows that the dataset supports three genuinely "
     "different approaches:")
bullet("Colour-histogram + HOG features with a linear/RBF SVM. Simple, fast, no GPU needed.",
       "Method 1 — traditional baseline: ")
bullet("A small convolutional network (about 4 conv layers) trained on the 64 × 64 images from random "
       "initialisation, with flip/rotation augmentation.", "Method 2 — small CNN from scratch: ")
bullet("ResNet-18 pre-trained on ImageNet, fine-tuned on EuroSAT.", "Method 3 — transfer learning: ")
para("Planned metrics: overall accuracy, macro precision, macro recall and macro F1-score, plus a "
     "confusion matrix for the analysis.")

# ---------- 6. Disclosure ----------
heading("6. Source and AI Disclosure", level=1)
caption("Table 4. External sources and AI tools used for this report")
table([
    ["Source / tool", "Used for", "What I changed or added myself"],
    ["EuroSAT dataset and paper [1], Zenodo record 7711810",
     "The data; class definitions; how the labels were made",
     "Dataset used as-is (no images changed); a stratified 70/15/15 split was added (Section 2c)"],
    ["Python libraries: NumPy, pandas, Pillow, Matplotlib (Jupyter)",
     "Loading images, counting, plotting", "—"],
    ["Claude (Anthropic), model Claude Opus 5.5, via the Claude Code desktop app",
     "Suggested the dataset; drafted the notebook code (download, counting, plots, annotation checks) "
     "and a first draft of this report's text",
     "I accepted the suggested dataset and the overall plan, reviewed the final report and the "
     "notebook output, and requested the final edits (author name, this disclosure). The notebook "
     "code and report text are otherwise used largely as drafted; the method code for the midterm "
     "will be written and decided by me."],
], widths=[1.9, 2.3, 2.3], font=9)
para("A case where the AI output was wrong:", bold=True)
para("The first version of the notebook's mean-RGB cell sampled 300 images per class with "
     "df.groupby('label', group_keys=False).apply(lambda g: g.sample(300)). When the notebook was run, "
     "this cell crashed with KeyError: 'label', because in recent versions of pandas groupby().apply() "
     "no longer passes the grouping column to the function, so the next line could not group by it. "
     "I noticed it because the notebook stopped at that cell and the figure file was not created. It "
     "was fixed by replacing the call with df.groupby('label').sample(300, random_state=0), which keeps "
     "the column; the notebook then ran from top to bottom without errors.")

# ---------- References ----------
heading("References", level=1)
para("[1] P. Helber, B. Bischke, A. Dengel, D. Borth. “EuroSAT: A Novel Dataset and Deep Learning Benchmark "
     "for Land Use and Land Cover Classification.” IEEE Journal of Selected Topics in Applied Earth "
     "Observations and Remote Sensing, 12(7):2217–2226, 2019.", size=10)
para("[2] P. Helber et al. “Introducing EuroSAT: A Novel Dataset and Deep Learning Benchmark for Land Use "
     "and Land Cover Classification.” IGARSS 2018.", size=10)
para("[3] EuroSAT RGB download: https://zenodo.org/records/7711810 ; project page: "
     "https://github.com/phelber/EuroSAT", size=10)

doc.save("Dataset_Selection_Report.docx")


def footer(canvas, d):
    canvas.saveState()
    canvas.setFont("Arial", 8.5)
    canvas.setFillColor(colors.grey)
    canvas.drawCentredString(letter[0] / 2, 0.5 * inch, f"Page {d.page}")
    canvas.restoreState()


def keep_headings(items):
    """Glue each heading (and any heading right after it) to the first block that follows."""
    out, buf = [], []
    for it in items:
        is_head = isinstance(it, Paragraph) and it.style.name in ("h1", "h2")
        if is_head:
            buf.append(it)
        elif buf:
            inner = it._content if isinstance(it, KeepTogether) else [it]   # flatten: nested KeepTogether breaks pages
            out.append(KeepTogether(buf + list(inner)))
            buf = []
        else:
            out.append(it)
    return out + buf


story = keep_headings(story)
SimpleDocTemplate("Dataset_Selection_Report.pdf", pagesize=letter, leftMargin=inch, rightMargin=inch,
                  topMargin=0.85 * inch, bottomMargin=0.85 * inch, title="Dataset Selection Report — EuroSAT",
                  author="").build(story, onFirstPage=footer, onLaterPages=footer)
print("saved")
