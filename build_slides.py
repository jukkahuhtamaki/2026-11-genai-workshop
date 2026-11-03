#!/usr/bin/env python3
# Builds the case template for the workshop: python build_slides.py -> case-template.pptx
#
# Slides 1-3: blank template with guiding questions
# Slides 4-6: the same slides filled in with an illustrative example case

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR as T
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

THEME = {
    "dk1": "17322E",      # forest ink
    "lt1": "FFFFFF",
    "dk2": "4D5F5B",      # muted text
    "lt2": "EEF3F1",      # card tint
    "accent1": "D9573B",  # coral
    "accent2": "2A7F78",  # teal
    "accent3": "E3A93B",  # amber
    "accent4": "17322E",
    "accent5": "CFE3DF",  # pale teal
    "accent6": "F7DCD4",  # pale coral
    "hlink": "2A7F78",
    "folHlink": "4D5F5B",
}
HEAD_FONT, BODY_FONT = "Cambria", "Calibri"

W, H = 13.333, 7.5
M = 0.6
GAP = 0.3
WORKSHOP = "Responsible use of GenAI in research · workshop"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)


# --- theme and layouts -----------------------------------------------------

def apply_theme():
    part = prs.slide_master.part.part_related_by(RT.THEME)
    tree = etree.fromstring(part.blob)
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    tree.find(f".//{a}clrScheme").set("name", "Workshop")
    for name, value in THEME.items():
        el = tree.find(f".//{a}clrScheme/{a}{name}")
        for child in list(el):
            el.remove(child)
        etree.SubElement(el, f"{a}srgbClr", val=value)
    fonts = tree.find(f".//{a}fontScheme")
    fonts.set("name", "Workshop")
    fonts.find(f"{a}majorFont/{a}latin").set("typeface", HEAD_FONT)
    fonts.find(f"{a}minorFont/{a}latin").set("typeface", BODY_FONT)
    part._blob = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


def setup_title_style():
    master = prs.slide_master
    lvl = master.element.find(qn("p:txStyles")).find(qn("p:titleStyle")).find(qn("a:lvl1pPr"))
    lvl.set("algn", "l")
    rpr = lvl.find(qn("a:defRPr"))
    rpr.set("sz", "3200")
    rpr.set("b", "1")
    for old in rpr.findall(qn("a:solidFill")):
        rpr.remove(old)
    fill = etree.Element(qn("a:solidFill"))
    etree.SubElement(fill, qn("a:schemeClr"), val="tx1")
    rpr.insert(0, fill)

    box = (Inches(M), Inches(0.75), Inches(W - 2 * M), Inches(0.9))
    for shapes in [master.placeholders] + [l.placeholders for l in prs.slide_layouts]:
        for ph in shapes:
            if ph.placeholder_format.type == 1:
                ph.left, ph.top, ph.width, ph.height = box
                ph.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                ph.text_frame.word_wrap = True


apply_theme()
setup_title_style()
TITLE_ONLY = prs.slide_layouts[5]


# --- helpers ---------------------------------------------------------------

def color(target, c):
    if isinstance(c, str):
        target.rgb = RGBColor.from_string(c)
    else:
        target.theme_color = c


def text(slide, x, y, w, h, paragraphs, size=14, col=T.TEXT_1, bold=False, italic=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, name=None, space_after=0, bullets=False):
    """paragraphs: a string, or a list where each item is a string or a list of
    (text, {options}) runs."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    if isinstance(paragraphs, str):
        paragraphs = [paragraphs]
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if space_after and i < len(paragraphs) - 1:
            p.space_after = Pt(space_after)
        if bullets:
            set_bullet(p)
        if isinstance(para, str):
            para = [(para, {})]
        for content, opts in para:
            r = p.add_run()
            r.text = content
            f = r.font
            f.size = Pt(opts.get("size", size))
            f.bold = opts.get("bold", bold)
            f.italic = opts.get("italic", italic)
            color(f.color, opts.get("col", col))
    return box


def set_bullet(p, indent=0.2):
    ppr = p._p.get_or_add_pPr()
    ppr.set("marL", str(int(Inches(indent))))
    ppr.set("indent", str(-int(Inches(indent))))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for el in ppr.findall(qn(tag)):
            ppr.remove(el)
    bu = etree.SubElement(ppr, qn("a:buChar"), char="•")
    return bu


def shape(slide, kind, x, y, w, h, fill=None, line=None, name=None, radius=None):
    s = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        color(s.fill.fore_color, fill)
    if line is None:
        s.line.fill.background()
    else:
        color(s.line.color, line)
        s.line.width = Pt(1.25)
    s.shadow.inherit = False
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    return s


def card(slide, x, y, w, h, fill=T.BACKGROUND_2, name=None, line=None):
    return shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, line, name, radius=0.06)


def pill(slide, x, y, w, h, label, fill=T.ACCENT_2, col=T.BACKGROUND_1, size=12,
         bold=True, line=None, name=None):
    s = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, line, name, radius=0.5)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    for i, line_ in enumerate(label.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = line_
        r.font.size = Pt(size)
        r.font.bold = bold
        color(r.font.color, col)
    return s


def icon_dot(slide, x, y, number, fill):
    shape(slide, MSO_SHAPE.OVAL, x, y, 0.42, 0.42, fill, name=f"Marker {number}")
    text(slide, x, y, 0.42, 0.42, str(number), size=14, bold=True, col=T.BACKGROUND_1,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def new_slide(title, step, example):
    slide = prs.slides.add_slide(TITLE_ONLY)
    slide.shapes.title.text = title
    label = f"{step} / 3   ·   {['THE CASE', 'THE SETUP', 'REFLECTION'][step - 1]}"
    text(slide, M, 0.42, 6, 0.3, label, size=12, bold=True, col=T.ACCENT_2, name="Step")
    if example:
        pill(slide, W - M - 2.6, 0.35, 2.6, 0.38, "Illustrative example", T.ACCENT_6,
             T.ACCENT_1, size=12, name="Example tag")
    text(slide, M, H - 0.38, 8, 0.25, WORKSHOP, size=10, col=T.TEXT_2, name="Footer")
    return slide


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


def grid_card(slide, i, title, body, cols=3, top=1.8, h=2.35, highlight=False,
              template=False, marker_fill=T.ACCENT_2):
    """A card in a 3 x 2 grid. body: list of bullet strings."""
    cw = (W - 2 * M - (cols - 1) * GAP) / cols
    x = M + (i % cols) * (cw + GAP)
    y = top + (i // cols) * (h + GAP)
    fill = T.TEXT_1 if highlight else T.BACKGROUND_2
    head = T.BACKGROUND_1 if highlight else T.TEXT_1
    body_col = T.ACCENT_5 if highlight else (T.TEXT_2 if template else T.TEXT_1)
    card(slide, x, y, cw, h, fill, name=f"Card {title}")
    icon_dot(slide, x + 0.25, y + 0.22, i + 1, T.ACCENT_1 if highlight else marker_fill)
    text(slide, x + 0.8, y + 0.2, cw - 1.0, 0.46, title, size=17, bold=True, col=head,
         anchor=MSO_ANCHOR.MIDDLE, name=f"Heading {title}")
    text(slide, x + 0.25, y + 0.8, cw - 0.5, h - 0.95, body, size=14, col=body_col,
         italic=template, bullets=True, space_after=4, name=f"Body {title}")


GAIDET = ["Conceptuali-\nzation", "Literature\nreview", "Methodology", "Software\ndevelopment",
          "Data\nmanagement", "Writing", "Ethical\noversight"]


def gaidet_strip(slide, used, template):
    y = 5.55
    text(slide, M, y - 0.45, 12, 0.35,
         [[("WHERE WAS GENAI USED?  ", {"bold": True, "col": T.TEXT_2}),
           ("Research stages from GAIDeT (Suchikova et al. 2025)"
            + (". Highlight the stages you used." if template else ""),
            {"col": T.TEXT_2})]], size=12)
    pw = (W - 2 * M - 6 * 0.12) / 7
    for i, stage in enumerate(GAIDET):
        on = i in used
        pill(slide, M + i * (pw + 0.12), y, pw, 0.72, stage,
             T.ACCENT_2 if on else T.BACKGROUND_1, T.BACKGROUND_1 if on else T.TEXT_2,
             size=12, bold=on, line=None if on else T.ACCENT_5,
             name=f"Stage {stage.replace(chr(10), '')}")


# --- slide builders --------------------------------------------------------

def case_slide(c, template):
    s = new_slide(c["title"], 1, not template)
    body_col = T.TEXT_2 if template else T.TEXT_1
    # left: question
    text(s, M, 1.95, 7.6, 0.3, "RESEARCH QUESTION", size=12, bold=True, col=T.TEXT_2)
    text(s, M, 2.3, 7.6, 1.2, c["question"], size=20, col=body_col, italic=template,
         name="Research question")
    text(s, M, 3.6, 7.6, 0.3, "WHAT WAS DONE", size=12, bold=True, col=T.TEXT_2)
    text(s, M, 3.95, 7.6, 1.0, c["done"], size=14, col=body_col, italic=template,
         name="What was done")
    # right: facts
    x, cw = 8.65, W - M - 8.65
    card(s, x, 1.95, cw, 3.0, name="Facts")
    rows = [("WHO", c["who"]), ("APPROACH", c["approach"]), ("OUTCOME", c["outcome"])]
    y = 2.15
    for label, value in rows:
        text(s, x + 0.25, y, cw - 0.5, 0.25, label, size=11, bold=True, col=T.ACCENT_2)
        text(s, x + 0.25, y + 0.27, cw - 0.5, 0.6, value, size=14, col=body_col,
             italic=template, name=f"Fact {label.title()}")
        y += 0.92
    gaidet_strip(s, c["stages"], template)
    notes(s, c["notes"])


def grid_slide(step, title, cards, template, highlight=None, slide_notes=""):
    s = new_slide(title, step, not template)
    for i, (heading, body) in enumerate(cards):
        grid_card(s, i, heading, body, highlight=(i == highlight), template=template)
    notes(s, slide_notes)


# --- content: template -----------------------------------------------------

TEMPLATE_CASE = {
    "title": "Case title: what did you do?",
    "question": "What did you want to find out or achieve?",
    "done": "In two or three sentences: what did you do, and what role did GenAI play?",
    "who": "Names, discipline, institution",
    "approach": "Qualitative, quantitative or mixed?",
    "outcome": "What came out of it, or what do you expect?",
    "stages": [],
    "notes": "Template slide 1/3. Replace the grey italic questions with your own text, "
             "and highlight the research stages where you used GenAI.",
}

TEMPLATE_SETUP = [
    ("Infrastructure", ["Where did the models run: laptop, university or national computing, cloud API?",
                        "Why there?"]),
    ("Tools and models", ["Which tools and models, which versions?", "When were they used?"]),
    ("Data", ["What data, from where?", "How sensitive? Does consent or the licence allow AI processing?"]),
    ("Methods", ["What was the workflow?", "What did you prompt, and how were outputs used?"]),
    ("Documentation", ["What did you log?", "Could someone check or repeat what you did?"]),
    ("Time and cost", ["How long did it take, compared with doing it without GenAI?",
                       "What did it cost?"]),
]

TEMPLATE_REFLECTION = [
    ("Barriers", ["Technical, ethical, legal, organisational or skills-related?"]),
    ("Opportunities", ["What did you discover that you did not expect?"]),
    ("Framework or canvas", ["Which framework or canvas helped you?", "What did it help with?"]),
    ("Epistemic authority", ["How did you check GenAI outputs?",
                             "Which decisions stayed with you?",
                             "What was GenAI not allowed to do?"]),
    ("Qualitative lens", ["What happened to interpretation and reflexivity?",
                          "Whose voices could be lost or changed?"]),
    ("New understanding", ["What do you now see differently about GenAI practices in research?"]),
]

# --- content: illustrative example -------------------------------------------

EXAMPLE_CASE = {
    "title": "Whose knowledge trains the models?",
    "question": "How prominent are Stack Exchange and other public archives in the "
                "training datasets shared on Hugging Face?",
    "done": "A crawler for Hugging Face dataset and model metadata was built by "
            "“academic vibe coding” with a GenAI coding assistant. The researchers specified "
            "and checked; the assistant wrote the code.",
    "who": "A research team studying open data and AI",
    "approach": "Mixed: metadata analysis + close reading of dataset cards",
    "outcome": "A reusable crawler and a dataset for comparing archives",
    "stages": [2, 3, 4],
    "notes": "Illustrative example: an imagined case, not real results.\n"
             "Archives compared: the Stack Exchange data dump against Wikipedia, arXiv, GitHub, "
             "Common Crawl and Reddit.\n"
             "“Academic vibe coding”: the researcher describes what the tool should do in plain "
             "language and a GenAI assistant writes the code.",
}

EXAMPLE_SETUP = [
    ("Infrastructure", ["Researcher’s laptop, coding assistant in a sandbox",
                        "Hugging Face Hub API, results in SQLite",
                        "Code in a public GitHub repository"]),
    ("Tools and models", ["Claude Code (Claude Opus 5.5) wrote the crawler",
                          "Python, huggingface_hub library",
                          "Tool versions and dates recorded"]),
    ("Data", ["Public metadata: downloads, likes, licence, source datasets",
              "Dataset cards as text, models’ declared training data",
              "No personal data; Stack Exchange is CC BY-SA"]),
    ("Methods", ["Each archive = a list of search terms",
                 "Prominence: datasets, downloads, models trained on them",
                 "Dataset cards coded for provenance"]),
    ("Documentation", ["Every row stores its query and retrieval time",
                       "Development log: what AI wrote, who checked it, how",
                       "Prompts kept in the repository"]),
    ("Time and cost", ["First working crawler in an afternoon",
                       "Estimate without GenAI: one to two weeks",
                       "API use: a few euros"]),
]

EXAMPLE_REFLECTION = [
    ("Barriers", ["Search matches names, not contents: mixtures like The Pile hide Stack Exchange",
                  "Metadata is self-reported and incomplete"]),
    ("Opportunities", ["Time moves from coding to research design",
                       "Easy to add new archives and rerun"]),
    ("Framework or canvas", ["GAIDeT to declare where AI was used",
                             "VT verification log, Datasheets for Datasets"]),
    ("Epistemic authority", ["Researchers set archives, measures and interpretations",
                             "Counts spot-checked against the Hugging Face website",
                             "AI never coded the dataset cards"]),
    ("Qualitative lens", ["Dataset cards read as texts: what do creators claim about consent?",
                          "Our choice of “archives” shapes what we find"]),
    ("New understanding", ["Vibe-coded code is a research instrument and needs the same scrutiny "
                           "as a survey"]),
]


# --- build -----------------------------------------------------------------

for template in (True, False):
    case = TEMPLATE_CASE if template else EXAMPLE_CASE
    setup = TEMPLATE_SETUP if template else EXAMPLE_SETUP
    reflection = TEMPLATE_REFLECTION if template else EXAMPLE_REFLECTION
    kind = "Template" if template else "Illustrative example"
    case_slide(case, template)
    grid_slide(2, "The setup: infrastructure, tools, data and methods" if template
               else "The setup: a crawler built by academic vibe coding", setup, template,
               slide_notes=f"{kind}, slide 2/3. Concrete details are the point: names, versions, "
                           "places, times.")
    grid_slide(3, "Reflection: barriers, opportunities and authority" if template
               else "Reflection: who knows what, and how?", reflection, template, highlight=3,
               slide_notes=f"{kind}, slide 3/3. Epistemic authority: how do you make sure the "
                           "researcher, not the tool, stays responsible for knowledge claims?\n"
                           "Background: Virginia Tech, Considerations for the Responsible and "
                           "Ethical Use of AI in Research (rev. Oct 14, 2026).")

prs.save("case-template.pptx")
print("case-template.pptx")
