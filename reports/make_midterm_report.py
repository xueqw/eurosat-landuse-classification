"""Builds reports/midterm/Midterm_Report.docx and .pdf. Run from the repository root after the notebooks."""
import os
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




S_EQ = ParagraphStyle("eq", fontName="Arial", fontSize=10.5, leading=15, alignment=TA_CENTER,
                      spaceBefore=2, spaceAfter=8)


def equation(markup):
    """Centred equation. `markup` may contain <sub>..</sub>; rendered as subscripts in PDF and Word."""
    import re
    lines = re.split(r"\s{4,}", markup)          # wide gaps separate two equations -> one per line
    if len(lines) > 1:
        for ln in lines:
            equation(ln)
        return
    story.append(Paragraph(markup, S_EQ))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for part in re.split(r"(<sub>.*?</sub>)", markup):
        if part.startswith("<sub>"):
            r = p.add_run(part[5:-6])
            r.font.subscript = True
        elif part:
            p.add_run(part)

# ---------- numbers come straight from the saved result files ----------
import json, subprocess
import pandas as pd
RES = "results/"
m1 = json.load(open(RES + "method1_test_metrics.json"))
search = pd.read_csv(RES + "method1_val_search.csv")
groups = pd.read_csv(RES + "method1_feature_groups_val.csv")
cls_report = open(RES + "method1_classification_report.txt").read()
REPO_URL = "https://github.com/xueqw/eurosat-landuse-classification"
# timestamp of the latest commit on the repository (the report is built right after pushing)
COMMIT = subprocess.run(["git", "log", "-1", "--format=%h|%cd", "--date=format:%Y-%m-%d %H:%M:%S %z"],
                        capture_output=True, text=True).stdout.strip().split("|")


def pct(x):
    return f"{100 * float(x):.2f}"


# ---------- repository line (must be at the very top of page 1) ----------
para(f"GitHub repository: {REPO_URL}", bold=True, size=10.5)
para(f"Latest commit: {COMMIT[0]} · {COMMIT[1]}", bold=True, size=10.5)

# ---------- title block ----------
p = para("CPS 4841 Computer Vision — Course Project", align=WD_ALIGN_PARAGRAPH.CENTER, size=16, bold=True)
p = para("Midterm Report: Land-Use Classification of Satellite Images (EuroSAT)",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=12.5)
para("Student: Qiwen Xue    ·    Section 03    ·    November 22, 2026",
     align=WD_ALIGN_PARAGRAPH.CENTER, size=10.5)

heading("0. Updated Dataset Description — What Changed", level=1)
para("Sections 1–4 repeat the dataset selection report. The dataset, the task and the problem definition are "
     "unchanged. Two practical updates: (1) the download and the fixed 70/15/15 split now live in common.py, "
     "so every method notebook loads exactly the same train/validation/test images (I checked that the "
     "regenerated split is identical to splits.csv for all 27,000 images); (2) the code is now in the "
     "GitHub repository above, with data excluded through .gitignore and a download script instead.")

# ---------- 1-4. Dataset (unchanged from the dataset selection report) ----------
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
caption("Table 1. Dataset overview")
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

# ---------- 5. Evaluation metrics ----------
heading("5. Evaluation Metrics", level=1)
para("All methods are evaluated on the same 4,050-image test split. Notation: N test images, true labels "
     "y(i), predicted labels ŷ(i), K = 10 classes. For one class k, TP(k) = images of class k predicted as k, "
     "FP(k) = images of other classes predicted as k, FN(k) = images of class k predicted as another class. "
     "All four metrics are computed with scikit-learn (common.compute_metrics) and range from 0 to 1 "
     "(reported as %; higher is better).")

heading("a. Accuracy", level=2)
equation("Accuracy = (1/N) · Σ<sub>i</sub> 1[ ŷ<sub>i</sub> = y<sub>i</sub> ]")
para("The fraction of test images whose predicted class is correct. It is the most common headline number "
     "for EuroSAT, so it lets me compare with published results. Its weakness is that larger classes "
     "count more, which is why the macro-averaged metrics below are also reported.")

heading("b. Macro-averaged precision", level=2)
equation("Precision<sub>k</sub> = TP<sub>k</sub> / (TP<sub>k</sub> + FP<sub>k</sub>)        Macro-Precision = (1/K) · Σ<sub>k</sub> Precision<sub>k</sub>")
para("Of the images the model labels as class k, how many really are class k. A low precision for a class "
     "means the model over-uses that label (for example calling many fields “PermanentCrop”). "
     "For a land-cover map this is the error a map user sees: the map says “Industrial”, but it is not.")

heading("c. Macro-averaged recall", level=2)
equation("Recall<sub>k</sub> = TP<sub>k</sub> / (TP<sub>k</sub> + FN<sub>k</sub>)        Macro-Recall = (1/K) · Σ<sub>k</sub> Recall<sub>k</sub>")
para("Of the images that truly belong to class k, how many the model finds. Low recall means a class is "
     "missed, e.g. rivers labelled as highways. Macro-averaged recall is the same as balanced accuracy, "
     "so it is not inflated by the larger classes.")

heading("d. Macro-averaged F1-score", level=2)
equation("F1<sub>k</sub> = 2 · Precision<sub>k</sub> · Recall<sub>k</sub> / (Precision<sub>k</sub> + Recall<sub>k</sub>)        Macro-F1 = (1/K) · Σ<sub>k</sub> F1<sub>k</sub>")
para("The harmonic mean of precision and recall for each class, averaged with equal weight over the ten "
     "classes. It is my main metric for model selection (e.g. choosing the SVM's C) because one number "
     "punishes both over-predicting and missing a class, and every class counts equally.")

heading("e. Why these metrics fit EuroSAT", level=2)
bullet("The classes are nearly but not perfectly balanced (2,000–3,000 images, ratio 1.5 : 1). Accuracy "
       "alone would weigh Pasture (7.4 %) less than Forest (11.1 %); macro averages treat them equally.",
       "Mild imbalance: ")
bullet("The interesting errors are between similar classes (crop types, vegetation types, roads vs. "
       "rivers). Per-class precision and recall, plus a confusion matrix, show exactly which pairs fail, "
       "which the final report's error analysis needs.", "Similar classes: ")
bullet("Accuracy is what the EuroSAT paper and most later work report, so my numbers can be put in "
       "context.", "Comparability: ")

# ---------- 6. Method 1 ----------
heading("6. Method 1: Hand-Crafted Features + SVM (Traditional Baseline)", level=1)
heading("a. Introduction", level=2)
para("Method 1 is the simple baseline required by the project. It does not learn features: each image is "
     "described by fixed, hand-designed features, and a classifier learns only how to combine them. "
     "I combined three classic descriptors because the dataset figures (Figure 2) showed that colour "
     "alone separates some classes (water, forest) but not others (highway vs. residential), which need "
     "edge and texture information:")
bullet("Histogram of Oriented Gradients [4]: the image is split into 8 × 8-pixel cells; each cell gets a "
       "9-bin histogram of gradient directions, normalised over 2 × 2-cell blocks. It describes edges "
       "and shapes such as straight roads, field borders and building outlines.", "HOG ")
bullet("Local Binary Patterns [5]: each pixel is compared with its 8 neighbours to form a binary code; "
       "the histogram of the 10 “uniform” codes describes fine texture (rough tree canopy vs. smooth "
       "water).", "LBP ")
bullet("16-bin histograms of hue, saturation and value, plus the mean and standard deviation of R, G "
       "and B. They describe colour and brightness.", "Colour ")
para("The concatenated vector (1,828 numbers) is standardised and classified by a support vector machine "
     "with an RBF kernel [6], which can learn non-linear class boundaries in this feature space.")

heading("b. Implementation details", level=2)
caption("Table 4. Method 1 implementation details")
table([
    ["Item", "Setting"],
    ["Code", "02_method1_hog_color_svm.ipynb (scikit-image 0.26, scikit-learn 1.9); helpers in common.py"],
    ["Input / preprocessing", "64 × 64 RGB, no resizing or augmentation; grayscale copy for HOG and LBP"],
    ["HOG", "9 orientations, 8 × 8 pixels per cell, 2 × 2 cells per block, L2-Hys → 1,764 values"],
    ["LBP", "uniform, P = 8 neighbours, radius 1, normalised 10-bin histogram → 10 values"],
    ["Colour", "HSV 3 × 16-bin histograms (sum to 1) + RGB mean and std / 255 → 54 values"],
    ["Feature scaling", "StandardScaler fitted on the training split only"],
    ["Classifier", "SVC, RBF kernel, gamma = 'scale' (1 / (n_features · Var(X))), one-vs-one multi-class"],
    ["Hyper-parameter search", "C = 1, 3, 10 or 30, each trained on train, compared on validation by macro F1"],
    ["Final model", f"C = {m1['C']:g}, trained on the 18,900 training images, evaluated once on the test split"],
    ["Random seed", "42 (split and SVM)"],
    ["Hardware", "Laptop CPU (Apple M5); no GPU needed"],
], widths=[1.7, 4.8])

# ---------- 7. Results ----------
heading("7. Results of Method 1", level=1)
heading("a. Test-set results", level=2)
caption("Table 5. Method 1 on the test split (4,050 images). Higher is better.")
table([
    ["Method", "Accuracy (%)", "Macro Precision (%)", "Macro Recall (%)", "Macro F1 (%)"],
    ["Method 1: HOG + LBP + colour, RBF SVM", pct(m1["accuracy"]), pct(m1["macro_precision"]),
     pct(m1["macro_recall"]), pct(m1["macro_f1"])],
    ["Method 2: small CNN from scratch", "final report", "", "", ""],
    ["Method 3: ResNet-18 transfer learning", "final report", "", "", ""],
], widths=[2.5, 1.0, 1.0, 1.0, 1.0])
para(f"Training took {m1['train_seconds']:.0f} s on a laptop CPU and prediction about "
     f"{m1['predict_ms_per_image']:.1f} ms per image; the final SVM keeps {m1['n_support_vectors']:,} "
     f"support vectors. All numbers are copied by the report script from results/method1_test_metrics.json, "
     f"which the notebook writes.")

heading("b. Choosing C on the validation split", level=2)
caption("Table 6. Validation results for each value of C (all features). The best macro F1 is used.")
rows = [["C", "Accuracy (%)", "Macro Precision (%)", "Macro Recall (%)", "Macro F1 (%)", "Fit + val. prediction (s)"]]
for _, r in search.iterrows():
    rows.append([f"{r['C']:g}" + (" (chosen)" if r["C"] == m1["C"] else ""), pct(r["accuracy"]),
                 pct(r["macro_precision"]), pct(r["macro_recall"]), pct(r["macro_f1"]), f"{r['fit_seconds']:.0f}"])
table(rows, widths=[1.0, 1.0, 1.15, 1.05, 1.0, 0.9])

heading("c. Per-class results and confusion matrix", level=2)
para(f"The final model misclassifies {round((1 - m1['accuracy']) * m1['n_test'])} of the {m1['n_test']:,} test images. "
     "Per-class results (scikit-learn classification_report, notebook output):")
code_block(cls_report.strip("\n"), size=8)
figure(RES + "method1_confusion_matrix.png", 4.9,
       "Figure 4. Method 1 test confusion matrix. Numbers are image counts; colour is the share of the true "
       "class (row).")

para(f"Overall, Method 1 classifies {pct(m1['accuracy'])} % of the test images correctly, with a macro F1 of "
     f"{pct(m1['macro_f1'])} %. The test numbers are close to the validation numbers for the chosen C "
     f"(accuracy {pct(search.loc[search['C'] == m1['C'], 'accuracy'].iloc[0])} %, macro F1 "
     f"{pct(search.loc[search['C'] == m1['C'], 'macro_f1'].iloc[0])} %), so choosing C on the validation split "
     "did not overfit it. For context, fine-tuned deep CNNs reach about 98 % accuracy on EuroSAT RGB [1], "
     "so there is clear room for Methods 2 and 3.")
bullet("SeaLake (F1 98.0 %), Forest (97.8 %), Industrial (94.3 %) and Residential (91.7 %). These classes "
       "have a distinctive colour (dark water, dark green canopy) or a distinctive texture (dense bright "
       "roofs), which the colour and LBP features capture directly.", "Easy classes: ")
bullet("PermanentCrop (F1 69.4 %, recall only 64.8 %), Highway (76.0 %) and HerbaceousVegetation (78.9 %).",
       "Hard classes: ")
bullet("PermanentCrop ↔ HerbaceousVegetation (42 + 32 test images), PermanentCrop → AnnualCrop (40), "
       "River ↔ Highway (33 + 27), Highway → AnnualCrop (29) and HerbaceousVegetation → Residential (28). "
       "The vegetation and crop classes share the same greens and browns; what separates an orchard from "
       "a meadow or a crop field is the spatial layout (rows of trees, parcel size), which at 10 m per pixel "
       "spans only a few pixels and is not described by histograms or by HOG's small 8 × 8 cells. Rivers "
       "and highways are both thin lines crossing the patch, so their edge-direction (HOG) features look "
       "alike, and colour only helps when the water is clearly visible.", "Most frequent confusions: ")

heading("d. Which features matter (validation split)", level=2)
caption("Table 7. Each feature group alone vs. all features (validation split, C = 3).")
rows = [["Features", "Dimensions", "Accuracy (%)", "Macro F1 (%)"]]
for _, r in groups.iterrows():
    rows.append([r["features"], f"{int(r['dims']):,}", pct(r["accuracy"]), pct(r["macro_f1"])])
table(rows, widths=[2.4, 1.2, 1.2, 1.2])
para("Colour alone (54 numbers) is the strongest single group, LBP texture with only 10 numbers is second, "
     "and HOG alone is the weakest even though it has 1,764 numbers. Combining all three adds about 3.6 "
     "points of macro F1 over colour alone. A likely reason the gain is not larger is that, after "
     "standardisation, HOG makes up 96 % of the dimensions and therefore dominates the RBF kernel's "
     "distance, drowning out the more useful colour and texture features. Reducing HOG with PCA or "
     "weighting the groups is something I may test for the final report.")
bullet(f"{m1['n_support_vectors']:,} of the 18,900 training images "
       f"({100 * m1['n_support_vectors'] / 18900:.0f} %) end up as support vectors, which shows that the "
       "classes overlap heavily in this feature space. It also makes prediction slow for such small images "
       f"(about {m1['predict_ms_per_image']:.0f} ms per image on CPU).", "Cost: ")

heading("e. A first look at the mistakes", level=2)
figure(RES + "method1_errors.png", 6.5,
       "Figure 5. Twelve random misclassified test images (true class → predicted class).")
para("Several mistakes involve patches that contain more than one land use, for example a motorway "
     "interchange next to buildings (Highway → PermanentCrop) or field parcels crossed by a road "
     "(PermanentCrop → Highway). Others are visually close classes such as harvested AnnualCrop fields vs. "
     "PermanentCrop. The full error analysis of at least 10 cases, grouped into failure patterns, will be "
     "done in the final report on the best of the three methods.")

AI_ERROR_CASE = (
    "In the first draft of the Method 1 notebook, the line that builds the train/validation/test masks was "
    "written as tr, va, te = (df['split'] == s).to_numpy(), ... although the variable s was never defined, "
    "followed by a second, redundant line that rebuilt the training mask. Running it would have stopped the "
    "notebook with a NameError right after the 30-second feature extraction. The mistake was caught when the "
    "draft was read through before the first run, and the two lines were replaced by one list comprehension, "
    "tr, va, te = [(df['split'] == s).to_numpy() for s in ('train', 'val', 'test')]. The notebook then ran "
    "from top to bottom without errors. (The dataset report describes an earlier case: a pandas groupby call "
    "that crashed with KeyError: 'label'.)")


# ---------- 8. Plan ----------
heading("8. Plan for Methods 2 and 3 (Final Report)", level=1)
bullet("A small convolutional network (about four conv–BN–ReLU–pool blocks and a linear head) trained "
       "from random initialisation on the 64 × 64 images with flip and 90° rotation augmentation, "
       "cross-entropy loss and Adam. It tests whether features learned from EuroSAT alone beat the "
       "hand-crafted ones.", "Method 2 — small CNN from scratch: ")
bullet("ResNet-18 pre-trained on ImageNet [7], last layer replaced by a 10-class head and fine-tuned on "
       "EuroSAT (images resized to 224 × 224, Colab GPU). It tests how well ImageNet features transfer to "
       "satellite images.", "Method 3 — transfer learning: ")
para("All three methods use the same split, the same four metrics and the same evaluation code "
     "(common.compute_metrics), so the final comparison table is directly comparable.")

# ---------- 9. Disclosure ----------
heading("9. Source and AI Disclosure", level=1)
caption("Table 8. External sources and AI tools used in this stage of the project")
table([
    ["Source / tool", "Used for", "What I changed or added myself"],
    ["EuroSAT dataset and paper [1], Zenodo record 7711810",
     "The data; class definitions; how the labels were made",
     "Dataset used as-is (no images changed); a stratified 70/15/15 split was added (Section 2c)"],
    ["HOG [4] and LBP [5] as implemented in scikit-image; SVM [6] and metrics in scikit-learn",
     "Method 1 features, classifier and evaluation",
     "Chose to combine HOG, LBP and colour features, their parameters, and tuning C on the validation split"],
    ["NumPy, pandas, Pillow, Matplotlib, Jupyter", "Data handling, plots, notebooks", "—"],
    ["Claude (Anthropic), model Claude Opus 5.5, via the Claude Code desktop app",
     "Suggested the dataset; drafted common.py, both notebooks, the README and a first draft of this "
     "report; created the GitHub repository and commits at my request",
     "I approved the dataset, the method plan and making the repository public, reviewed the code, "
     "results and report, and requested the final edits. Code drafted by the AI is marked with "
     "comments in the notebooks and in common.py."],
], widths=[1.9, 2.3, 2.3], font=9)
para("A case where the AI output was wrong:", bold=True)
para(AI_ERROR_CASE)

# ---------- References ----------
heading("References", level=1)
for ref in [
    "[1] P. Helber, B. Bischke, A. Dengel, D. Borth. “EuroSAT: A Novel Dataset and Deep Learning Benchmark "
    "for Land Use and Land Cover Classification.” IEEE JSTARS, 12(7):2217–2226, 2019.",
    "[2] P. Helber et al. “Introducing EuroSAT: A Novel Dataset and Deep Learning Benchmark for Land Use "
    "and Land Cover Classification.” IGARSS 2018.",
    "[3] EuroSAT RGB download: https://zenodo.org/records/7711810 ; project page: https://github.com/phelber/EuroSAT",
    "[4] N. Dalal, B. Triggs. “Histograms of Oriented Gradients for Human Detection.” CVPR 2005.",
    "[5] T. Ojala, M. Pietikäinen, T. Mäenpää. “Multiresolution Gray-Scale and Rotation Invariant Texture "
    "Classification with Local Binary Patterns.” IEEE TPAMI, 24(7):971–987, 2002.",
    "[6] C. Cortes, V. Vapnik. “Support-Vector Networks.” Machine Learning, 20:273–297, 1995.",
    "[7] K. He, X. Zhang, S. Ren, J. Sun. “Deep Residual Learning for Image Recognition.” CVPR 2016.",
]:
    para(ref, size=10)

os.makedirs("reports/midterm", exist_ok=True)
doc.save("reports/midterm/Midterm_Report.docx")


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
SimpleDocTemplate("reports/midterm/Midterm_Report.pdf", pagesize=letter, leftMargin=inch, rightMargin=inch,
                  topMargin=0.85 * inch, bottomMargin=0.85 * inch, title="Midterm Report — EuroSAT",
                  author="").build(story, onFirstPage=footer, onLaterPages=footer)
print("saved")
