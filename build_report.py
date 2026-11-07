"""Build the COM7024 Word submission from results.json and the official cover page."""

from __future__ import annotations

import json
import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from lxml import etree

ROOT = Path(__file__).resolve().parent
COVER = ROOT.parent / "Mathematics for DS" / "Arden_Assessment_Cover_page.docx"
RESULTS = ROOT / "results.json"
CODE = ROOT / "com7024_analysis.py"
OUT = ROOT / "COM7024_Programming_for_Data_Analytics_25106724.docx"
NAVY = RGBColor(0x1F, 0x4E, 0x79)


def money(value: float) -> str:
    return f"{value:,.0f}"


def p_text(value: float) -> str:
    if value < 0.001:
        return "p < 0.001"
    return f"p = {value:.3f}"


def set_run_font(run, name: str = "Calibri", size: int = 12, bold: bool = False, italic: bool = False, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def add_text(paragraph, text, **kwargs):
    run = paragraph.add_run(text)
    set_run_font(run, **kwargs)
    return run


def body_paragraph(doc: Document, text: str, *, size: int = 12, space_after: int = 8):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(space_after)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.15
    add_text(paragraph, text, size=size)
    return paragraph


def heading(doc: Document, text: str, level: int = 1):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.1
    add_text(paragraph, text, size=16 if level == 1 else 13, bold=True, color=NAVY)
    return paragraph


def caption(doc: Document, text: str):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(10)
    add_text(paragraph, text, size=10, italic=True, color=RGBColor(0x33, 0x33, 0x33))
    return paragraph


def add_picture(doc: Document, path: Path, width: float = 6.1):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(2)
    run = paragraph.add_run()
    run.add_picture(str(path), width=Inches(width))
    return paragraph


def shade(cell, hex_color: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def add_table(doc: Document, headers: list[str], rows: list[list[str]]):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        cell.text = ""
        add_text(cell.paragraphs[0], header, size=9, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        shade(cell, "1F4E79")
    for row in rows:
        cells = table.add_row().cells
        for index, value in enumerate(row):
            cells[index].text = ""
            add_text(cells[index].paragraphs[0], value, size=9)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(4)
    return table


def fix_running_header(doc: Document) -> None:
    """Keep the Arden logo in the header, clear of the body text."""
    section = doc.sections[0]
    section.top_margin = Inches(1.55)
    section.header_distance = Inches(0.28)
    section.bottom_margin = Inches(0.85)
    wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    drawing = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    header = section.header
    column_width = int(section.page_width - section.left_margin - section.right_margin)
    for anchor in header._element.iter(f"{wp}anchor"):
        anchor.set("behindDoc", "0")
        anchor.set("allowOverlap", "0")
        extent = anchor.find(f"{wp}extent")
        cx = 620000
        cy = int(620000 * 1095831 / 990600)
        if extent is not None:
            extent.set("cx", str(cx))
            extent.set("cy", str(cy))
        for ext in anchor.iter(f"{drawing}ext"):
            ext.set("cx", str(cx))
            ext.set("cy", str(cy))
        pos = anchor.find(f"{wp}positionH/{wp}posOffset")
        if pos is not None:
            pos.text = str(max(0, column_width - cx - 60000))
        vpos = anchor.find(f"{wp}positionV/{wp}posOffset")
        if vpos is not None:
            vpos.text = "0"
        for child in list(anchor):
            if child.tag.startswith(f"{wp}wrap"):
                anchor.remove(child)
        wrap = etree.Element(f"{wp}wrapSquare")
        wrap.set("wrapText", "left")
        effect = anchor.find(f"{wp}effectExtent")
        if effect is not None:
            effect.addnext(wrap)
        else:
            anchor.append(wrap)
    if len(header.paragraphs) < 2:
        header.add_paragraph()
    line = header.paragraphs[1]
    line.clear()
    line.alignment = WD_ALIGN_PARAGRAPH.LEFT
    line.paragraph_format.space_before = Pt(0)
    line.paragraph_format.space_after = Pt(0)
    add_text(line, "Student 25106724   ·   COM7024 Programming for Data Analytics", size=9, color=NAVY)


def fill_cover(doc: Document, word_count: int):
    answers = [
        "Arden University",
        "MSc Data Science",
        "Programming for Data Analytics (COM7024)",
        "Project",
        "25106724",
        "",
        f"{word_count:,} (main text only; cover, tables, figures, references and the code appendix are excluded)",
        "",
    ]
    for row, answer in zip(doc.tables[0].rows, answers):
        if not answer:
            continue
        paragraph = row.cells[0].add_paragraph()
        paragraph.paragraph_format.space_before = Pt(6)
        add_text(paragraph, answer, size=14, bold=True, color=NAVY)


def word_count(paragraphs) -> int:
    return sum(len(paragraph.text.split()) for paragraph in paragraphs)


def main():
    data = json.loads(RESULTS.read_text(encoding="utf-8"))
    audit = data["audit"]
    before = data["profile_before"]
    after = data["profile_after"]
    water = data["waterfront"]
    floors = data["floor_space"]["with_price"]
    by_group = data["floor_space"]["living_by_waterfront"]
    built = data["built_price"]
    model = data["model"]
    dist = data["distributions"]
    living_coef = model["coefficients_log"]["sqft_living"]
    per_100 = (pow(2.718281828, living_coef * 100) - 1) * 100

    shutil.copy(COVER, OUT)
    doc = Document(OUT)
    doc.add_page_break()

    title = doc.add_paragraph()
    title.paragraph_format.space_after = Pt(2)
    add_text(title, "Manchester housing analytics for Estate Management Plc", size=20, bold=True, color=NAVY)
    for line in (
        "Programming for Data Analytics (COM7024)",
        "Student number: 25106724",
        "Email: 25106724@ardenuniversity.ac.uk",
        "Assignment: Project",
    ):
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.space_before = Pt(0)
        add_text(paragraph, line, size=12)

    start = len(doc.paragraphs)

    heading(doc, "1. Purpose of the study")
    body_paragraph(
        doc,
        "Estate Management Plc asked for a Python analysis of the supplied file, Manchester_house_Dataset[3678].csv. "
        "The brief concentrates on waterfront properties, on the floor-space variables, and on the correlation between "
        "year built and price. The business purpose is a short, evidence-based view that can inform later thinking "
        "about the housing market. This report records the programming approach, the two statistical profiles required "
        "by the brief, the focused investigations, and a reflective note for the Senior Research Manager.",
    )
    body_paragraph(
        doc,
        "The cycle is deliberately small. The file is imported, types are corrected, every variable is profiled, rows "
        "that cannot represent a usable sale are removed, the profile is repeated, and the three questions are then "
        "tested with graphics and rank-based statistics. A log-linear model is used only to separate effects that "
        "overlap. It is a screen, not a valuation tool.",
    )

    heading(doc, "2. Programming approach")
    body_paragraph(
        doc,
        "The whole study is one script, com7024_analysis.py, so every table and figure in this document can be "
        "regenerated from the CSV. Three programming paradigms are used because each one fits a different part of the work. "
        "The top-level workflow is procedural: load, type, clean, test, plot and save. The frames and those steps are "
        "held on a HousingStudy object, so the raw file, the typed file and the cleaned file cannot be mixed up by a "
        "later plot. New columns, including price per square foot and decade of construction, are vectorised expressions "
        "on whole columns. That is the functional style of Pandas and NumPy, and it avoids a Python loop across nearly "
        "20,000 rows.",
    )
    body_paragraph(
        doc,
        "Types are set explicitly, which is the first analytical decision rather than a technicality. On import, id "
        "arrived as an integer. An identifier is not a quantity, so it is stored as a string and is kept out of every "
        "statistical profile. waterfront is stored as a nullable integer so a failed parse cannot become a silent float. "
        "In renovated, 0 means that no renovation year was recorded. A separate flag stores that meaning, so the zero is "
        "not treated as the year zero. Layout and names follow PEP 8 (van Rossum, Warsaw and Coghlan, 2001).",
    )
    body_paragraph(
        doc,
        "The libraries are the toolkit practised on this module. Pandas holds the tables (McKinney, 2010). NumPy supports "
        "the numerical steps (Harris et al., 2020). Matplotlib and Seaborn draw the figures (Hunter, 2007; Waskom, 2021). "
        "SciPy supplies the normality test, the rank correlations and the Mann-Whitney test (Virtanen et al., 2020). "
        "scikit-learn fits the held-out linear baseline (Pedregosa et al., 2011). The split uses student number 25106724 "
        "as the random seed, so the test-set result can be repeated. The script is published so the marker can rerun it: "
        "https://github.com/shaheer4290/com7024-manchester-housing-analysis",
    )
    body_paragraph(
        doc,
        "iLearn for this module could not be opened from outside a student login. The techniques below are the ones named "
        "in the assessment brief, in Arden’s published description of Programming for Data Analytics, and in Arden’s Python "
        "for data mining workshops: paradigms, Python fundamentals, PEP 8, files and modules, NumPy, Pandas, missing data "
        "and preprocessing, descriptive statistics, distributional checks, correlation, hypothesis tests, and Matplotlib "
        "and Seaborn graphics. Table A maps each of those topics to the part of the script that uses it.",
    )
    caption(doc, "Table A. Module techniques and where the script uses them.")
    add_table(
        doc,
        ["Technique", "Where it is used"],
        [
            ["Procedural workflow", "main(): load, type, clean, test, plot, save"],
            ["Object-oriented design", "class HousingStudy keeps raw, typed and cleaned frames apart"],
            ["Functions and return values", "waterfront_tests, floor_space_tests, built_price_tests, baseline_model"],
            ["Conditionals and iteration", "taught_patterns loops over decades and branches on the median"],
            ["Exception handling", "HousingStudy.load raises a clear error if the CSV is missing"],
            ["PEP 8 layout", "names, spacing and docstrings follow PEP 8"],
            ["File handling", "read_csv on the module file; results written as JSON and PNG"],
            ["Modules and libraries", "pandas, numpy, matplotlib, seaborn, scipy, sklearn"],
            ["NumPy arrays", "log price, train-test arrays, Kruskal groups"],
            ["Pandas frames", "selection, boolean filters, groupby, derived columns"],
            ["Types and missing values", "id as string; 13 missing bedrooms and 10 missing bathrooms removed"],
            ["Pre-processing", "duplicate ids dropped; impossible bedroom counts removed; price tail kept"],
            ["Descriptive statistics", "Tables 1 and 2, before and after cleaning"],
            ["Normality and skew", "D’Agostino-Pearson test and skewness on price"],
            ["Correlation", "Pearson and Spearman for floor space and year built"],
            ["Hypothesis tests", "Mann-Whitney on waterfront; Welch t-test on log price as a check; Kruskal-Wallis across decades"],
            ["Charts", "histogram, box plot, scatter, heatmap, line, bar"],
            ["Held-out model", "LinearRegression on log price, seed 25106724"],
        ],
    )

    heading(doc, "3. Import and the first statistical profile")
    body_paragraph(
        doc,
        f"The import contains {audit['rows_imported']:,} data rows. An unnamed index column was dropped, leaving 14 fields: "
        "id, price, bedrooms, bathrooms, sqft_living, sqft_total, floors, waterfront, view, condition, grade, built, "
        "renovated and living_area. Pandas read every field as numeric. Bedrooms and bathrooms were already floating-point "
        "values, which fits bathroom counts that include halves. Thirteen bedroom values and ten bathroom values were "
        "missing. Price, both floor-space fields, floors, waterfront, view, condition, grade, built, renovated and "
        "living_area had no missing values.",
    )
    body_paragraph(
        doc,
        f"Table 1 is the profile of the typed file, before any row was removed. Price is strongly right-skewed "
        f"(skew {before['price']['skew']:.2f}). The mean ({money(before['price']['mean'])}) sits above the median "
        f"({money(before['price']['median'])}), and the maximum is {money(before['price']['max'])}. A D'Agostino-Pearson "
        f"test on a random sample of {dist['before']['n_tested']:,} prices rejects normality "
        f"(K² = {dist['before']['dagostino_k2']:,.0f}, {p_text(dist['before']['p_value'])}). Means and a two-sample "
        "t-test would be a poor summary of this file. Later price comparisons therefore use medians, Spearman correlation "
        "and the Mann-Whitney U test. Waterfront is rare. renovated has a median of 0, so the typical row has no "
        f"renovation year. sqft_total is extremely skewed (skew {before['sqft_total']['skew']:.1f}), which is what plot "
        f"size looks like when a few parcels are very large. built runs from {before['built']['min']:.0f} to "
        f"{before['built']['max']:.0f}. The file does not name a currency. All prices below are in the units recorded "
        "in the file.",
    )
    caption(doc, "Table 1. First statistical profile, after type correction and before row removal.")
    variables = [
        "price", "bedrooms", "bathrooms", "sqft_living", "sqft_total", "floors",
        "waterfront", "view", "condition", "grade", "built", "renovated", "living_area",
    ]
    table_rows = []
    for name in variables:
        item = before[name]
        median = f"{item['median']:,.1f}" if name in {"bathrooms", "floors"} else f"{item['median']:,.0f}"
        table_rows.append([
            name,
            str(item["missing"]),
            median,
            f"{item['mean']:,.1f}",
            f"{item['std']:,.1f}",
            f"{item['skew']:.2f}",
            f"{item['min']:,.0f}",
            f"{item['max']:,.0f}",
        ])
    add_table(doc, ["Variable", "Missing", "Median", "Mean", "Std. dev.", "Skew", "Min", "Max"], table_rows)
    add_picture(doc, ROOT / "figures" / "fig0_execution_evidence.png", 6.2)
    caption(doc, "Figure 1. Execution evidence from com7024_analysis.py (student 25106724).")

    heading(doc, "4. Pre-processing and the second profile")
    body_paragraph(
        doc,
        f"Cleaning was limited to rows that cannot support a property comparison. {audit['duplicate_ids']} extra rows "
        f"shared a property id with an earlier row, and {audit['exact_duplicate_rows']} of those repeats were exact copies "
        "of another row. The file has no sale date, so a repeated id cannot be ordered in time. The first row for each id "
        f"was kept. After the repeats were collapsed, {audit['invalid_rows_removed']} rows still failed validation because "
        "bedrooms or bathrooms were missing, or the bedroom count was above 15. The maximum bedroom count on the raw file "
        f"was {before['bedrooms']['max']:.0f}, which is not a credible dwelling size beside the rest of this file. "
        f"Together, {audit['rows_removed']} rows were removed and {audit['rows_after_cleaning']:,} rows remained.",
    )
    body_paragraph(
        doc,
        "Price was not capped. Waterfront sales sit in the right-hand tail, and trimming high prices would have removed "
        "the contrast the estate manager asked about. The validation rule in HousingStudy.preprocess is:",
    )
    code = doc.add_paragraph()
    code.paragraph_format.space_after = Pt(8)
    code.paragraph_format.left_indent = Cm(0.75)
    add_text(
        code,
        "invalid where price, living area or plot size is missing or not positive;\n"
        "bedrooms are missing, not positive, or above 15; bathrooms are missing or not positive;\n"
        "year built is outside 1800-2025; or waterfront is not 0 or 1.",
        name="Consolas",
        size=9,
    )
    body_paragraph(
        doc,
        f"Table 2 and Figure 2 show the second profile. The median price is unchanged at {money(after['price']['median'])}. "
        f"Price skew stays at {after['price']['skew']:.2f}, and a second normality test on {dist['after']['n_tested']:,} "
        f"prices still rejects normality (K² = {dist['after']['dagostino_k2']:,.0f}, {p_text(dist['after']['p_value'])}). "
        "The gain from preprocessing is a file without duplicate identities, without missing room counts, and without the "
        f"extreme bedroom value. Bedroom skew falls from {before['bedrooms']['skew']:.2f} to {after['bedrooms']['skew']:.2f} "
        f"because that extreme had been dominating the moment. The minimum grade rises from {before['grade']['min']:.0f} "
        "to "
        f"{after['grade']['min']:.0f}. All tests in the rest of this report use the cleaned file.",
    )
    caption(doc, "Table 2. Selected statistics before and after row removal.")
    changed = []
    for name in variables:
        left, right = before[name], after[name]
        changed.append([
            name,
            f"{left['median']:,.0f}",
            f"{right['median']:,.0f}",
            f"{left['skew']:.2f}",
            f"{right['skew']:.2f}",
            f"{left['max']:,.0f}",
            f"{right['max']:,.0f}",
        ])
    add_table(
        doc,
        ["Variable", "Median before", "Median after", "Skew before", "Skew after", "Max before", "Max after"],
        changed,
    )
    add_picture(doc, ROOT / "figures" / "fig1_price_before_after.png", 6.3)
    caption(doc, "Figure 2. Price before cleaning and after invalid rows were removed. The tail remains.")

    heading(doc, "5. Waterfront properties")
    body_paragraph(
        doc,
        f"The cleaned file contains {water['n_waterfront']} waterfront sales, "
        f"{water['share_waterfront'] * 100:.2f}% of {audit['rows_after_cleaning']:,} rows, and "
        f"{water['n_inland']:,} inland sales. The waterfront median price is {money(water['median_price_waterfront'])}, "
        f"against {money(water['median_price_inland'])} inland, a ratio of {water['median_ratio']:.2f}. The means point "
        f"the same way ({money(water['mean_price_waterfront'])} against {money(water['mean_price_inland'])}), but the "
        "tail pulls them, so the median is the figure to use in a briefing. Waterfront homes are also larger: median "
        f"living space is {water['median_living_waterfront']:,.0f} square feet against "
        f"{water['median_living_inland']:,.0f}. The premium is not only size. Median price per square foot of living "
        f"space is {water['median_psf_waterfront']:,.0f} against {water['median_psf_inland']:,.0f}.",
    )
    body_paragraph(
        doc,
        f"A two-sided Mann-Whitney U test on price gives U = {water['mannwhitney_price_U']:,.0f} and "
        f"{p_text(water['mannwhitney_price_p'])}. The rank-biserial correlation is {water['rank_biserial']:.2f}, using the "
        "sign that is positive when waterfront prices occupy the higher ranks (Mann and Whitney, 1947). That is a large "
        f"separation. The same test on price per square foot is also {p_text(water['mannwhitney_psf_p'])}. Figure 3 hides "
        "the most extreme points so the quartiles can be read; the test itself uses every cleaned row. With 159 "
        "waterfront sales the premium is clear, and the group is too small to slice into many further categories. "
        f"A Welch t-test on log price agrees on the direction (t = {data['comparisons']['welch_log_price_t']:,.1f}, "
        f"{p_text(data['comparisons']['welch_log_price_p'])}). It is reported only as a check. Raw price is not normal, "
        "so the Mann-Whitney result remains the one used in the briefing.",
    )
    add_picture(doc, ROOT / "figures" / "fig2_waterfront_price.png", 5.6)
    caption(doc, "Figure 3. Sale price by waterfront status. Extreme points are hidden; the test uses all cleaned rows.")

    heading(doc, "6. Floor space")
    body_paragraph(
        doc,
        f"Internal living space is the floor-space field that tracks price. Pearson r between sqft_living and price is "
        f"{floors['sqft_living']['pearson_r']:.2f}, and Spearman r is {floors['sqft_living']['spearman_r']:.2f} "
        f"({p_text(floors['sqft_living']['spearman_p'])}, n = {floors['sqft_living']['n']:,}). The rank correlation is the "
        "better summary because price is skewed, and it is still strong. The relationship is tighter for waterfront homes "
        f"(Spearman r = {by_group['waterfront']['spearman_r']:.2f}, n = {by_group['waterfront']['n']}) than for inland homes "
        f"(Spearman r = {by_group['inland']['spearman_r']:.2f}). Figure 4 shows waterfront sales sitting higher and further "
        "to the right of the cloud.",
    )
    body_paragraph(
        doc,
        f"living_area also tracks price (Spearman r = {floors['living_area']['spearman_r']:.2f}), but it is not the same "
        "column as sqft_living. Until a data dictionary arrives it should be treated as a second size field, not as another "
        f"copy of the property's own floor area. Plot size is a weak associate of price (Spearman r = "
        f"{floors['sqft_total']['spearman_r']:.2f}, {p_text(floors['sqft_total']['spearman_p'])}). In this file, price "
        f"follows the building rather than the parcel. Floors have a moderate rank correlation with price "
        f"(Spearman r = {floors['floors']['spearman_r']:.2f}). Figure 6 shows price per square foot across floor counts, "
        "with wide uncertainty where a floor value is uncommon, so a difference of half a storey should not be briefed as "
        f"a pricing rule. Grade was outside the three named questions, and it correlates with price at Spearman r = "
        f"{floors['grade']['spearman_r']:.2f}. Any comment on size that ignores grade will mix quality into the size effect. "
        "Figure 5 puts these rank correlations on one grid.",
    )
    add_picture(doc, ROOT / "figures" / "fig3_floorspace_price.png", 5.8)
    caption(doc, "Figure 4. Internal floor space against price, with waterfront sales marked.")
    add_picture(doc, ROOT / "figures" / "fig4_correlation_heatmap.png", 5.5)
    caption(doc, "Figure 5. Spearman correlations among price, size, year built, waterfront and grade.")
    add_picture(doc, ROOT / "figures" / "fig6_floors_intensity.png", 5.6)
    caption(doc, "Figure 6. Price per square foot of living space by number of floors, with 95% intervals.")

    heading(doc, "7. Year built and price")
    body_paragraph(
        doc,
        f"The correlation requested in the brief is weak. Pearson r between built and price is "
        f"{built['overall']['pearson_r']:.3f} and Spearman r is {built['overall']['spearman_r']:.3f} "
        f"({p_text(built['overall']['spearman_p'])}, n = {built['overall']['n']:,}). The p-value is small because the "
        "sample is large. A rank correlation of 0.08 is a small commercial effect: construction year, on its own, "
        "organises very little of the price pattern. Figure 7 shows a wide cloud in every decade and a shallow line of "
        "decade medians.",
    )
    body_paragraph(
        doc,
        "Table 3 reads that line in numbers. Median price is higher for the oldest decade in the file than for the "
        "post-war decades, then rises again for the newest stock. Living space rises more steadily across the century. "
        "The newer stock is larger, so a simple statement that newer homes cost more mixes age with size. Within "
        f"waterfront sales the rank correlation is stronger (Spearman r = {built['by_waterfront']['waterfront']['spearman_r']:.2f}, "
        f"n = {built['by_waterfront']['waterfront']['n']}, {p_text(built['by_waterfront']['waterfront']['spearman_p'])}). "
        "That group has 159 rows and should be treated as a clue. Inland, the rank correlation remains small "
        f"(Spearman r = {built['by_waterfront']['inland']['spearman_r']:.3f}). There is no sale date. built is the "
        "construction year, so none of these figures is a trend in sale prices through calendar time. "
        f"A Kruskal-Wallis test does find a difference in price across the {data['comparisons']['n_decade_groups']} decades "
        f"(H = {data['comparisons']['kruskal_decade_H']:,.1f}, {p_text(data['comparisons']['kruskal_decade_p'])}). "
        "With this many sales a small shift is enough to reject the null, which matches the shallow median line in "
        "Figure 7 and the bars in Figure 8 rather than a strong age effect.",
    )
    caption(doc, "Table 3. Median price and median living space by decade of construction.")
    decade_rows = []
    for row in built["by_decade"]:
        decade_rows.append([
            str(int(row["decade"])),
            f"{int(row['n']):,}",
            money(row["median_price"]),
            f"{row['median_living']:,.0f}",
            f"{row['waterfront_share'] * 100:.1f}%",
        ])
    add_table(doc, ["Decade", "Sales", "Median price", "Median sqft living", "Waterfront share"], decade_rows)
    add_picture(doc, ROOT / "figures" / "fig5_built_price.png", 6.0)
    caption(doc, "Figure 7. Year built against price. The line is the median price of each decade, plotted at mid-decade.")
    add_picture(doc, ROOT / "figures" / "fig7_decade_bars.png", 6.0)
    caption(doc, "Figure 8. Median sale price by decade of construction.")

    heading(doc, "8. Separating overlapping effects")
    effects = model["approx_percent_effect"]
    body_paragraph(
        doc,
        f"A linear model of log price was fitted so that living space, plot size, floors, waterfront, year built, grade, "
        f"condition and view could be read together. The split was 75/25, giving {model['n_train']:,} training rows and "
        f"{model['n_test']:,} test rows. On the test set, R² on log price is {model['r2_log_price']:.2f}. After predictions "
        f"are transformed back to the recorded price, the mean absolute error is {money(model['mae_currency'])}. That error "
        "is small enough to show that the features carry information, and far too wide for a formal valuation.",
    )
    body_paragraph(
        doc,
        f"Holding the other features fixed, the approximate associations are about {per_100:.1f}% higher price for each "
        f"extra 100 square feet of living space, about {effects['floors']:.0f}% for an extra floor, about "
        f"{effects['waterfront']:.0f}% for the waterfront flag, and about {effects['grade']:.0f}% for an extra grade point. "
        f"Year built is slightly negative at about {effects['built']:.1f}% per year once size and grade are held constant. "
        f"Plot size rounds to {effects['sqft_total']:.0f}%. The raw waterfront median ratio of {water['median_ratio']:.2f} "
        "is therefore partly a size and quality difference. The flag still carries a large association after those are "
        "allowed for, and it remains a scarce feature. Grade is the strongest single lever in the model. A briefing that "
        "mentions only waterfront, floor space and age would miss the field that organises price most clearly. "
        f"Renovation is recorded on {dist['renovated_share'] * 100:.1f}% of cleaned rows. The median price on those rows is "
        f"{money(dist['median_price_renovated'])}, against {money(dist['median_price_never_renovated'])} where no renovation "
        "year is recorded. The file does not say whether the works caused the difference.",
    )

    heading(doc, "9. Reflective report for the Senior Research Manager")
    heading(doc, "Methodology", 2)
    body_paragraph(
        doc,
        "I treated the request as a decision brief. The script separates import, typing, cleaning and testing, writes "
        "the audit behind the tables, and saves the figures, so the same steps can be rerun when a corrected file arrives. "
        "Descriptive statistics were produced twice, because the brief asks for evidence that preprocessing changed the "
        "file. Price comparisons use medians and rank tests after the normality result. The log-linear model was added "
        "to see which associations remain when they are estimated together. Missing room counts were not imputed. Those "
        "rows are few, and an imputed bedroom count would have invented a layout.",
    )
    heading(doc, "Key findings", 2)
    body_paragraph(
        doc,
        "Waterfront sales are uncommon and much more expensive, in total price and in price per square foot. Internal "
        "floor space is the size measure that matters; plot size barely does. Year built and price have a detectable and "
        "commercially small correlation. Newer homes in the file are larger, and once size and grade are held constant "
        "the year coefficient changes sign. Reading the decade medians on their own would overstate the value of newer "
        "stock. Grade deserves a place in any later briefing even though it was outside the three named questions.",
    )
    heading(doc, "Limitations", 2)
    body_paragraph(
        doc,
        "The file is named as a Manchester housing dataset. Its fields are square feet, a grade running up to 13, a rare "
        "waterfront flag, and no postcode, tenure or sale date. The first record, id 7129300520 with a price of 221900, "
        "matches the opening row of a public King County, Washington, teaching file that is widely reused in data-science "
        "courses (Kaggle, n.d.). I cannot recommend that Estate Management Plc treats these numbers as Manchester evidence. "
        "If the file is that American extract under a new name, the recorded currency is the dollar, the water is not the "
        "Manchester Ship Canal, and the grade is a local building grade. The programming methods would still be the right "
        "methods for a UK file of the same shape. The market conclusions would not transfer.",
    )
    body_paragraph(
        doc,
        "Without a sale date there is no price index and no forecast. Future market trends cannot be estimated from "
        "construction year. Waterfront, size and grade are also chosen together by the people who build and buy homes. "
        "The waterfront coefficient is an association in a linear log-price model, not an estimate of what would happen "
        "if a house were moved onto the water. Keeping the first row when an id is repeated is a rule made necessary by "
        "the missing sale date. It is not evidence that the kept row is the latest sale.",
    )
    heading(doc, "Recommendations", 2)
    body_paragraph(
        doc,
        "Manchester asking prices, yields and waterside strategy should not be set from this file until its origin, "
        "geography, currency and sale date are confirmed in writing. The data owner should be asked for a data dictionary. "
        "If the file was supplied as a methods sample, it should be used for that purpose, and a Manchester extract should "
        "be commissioned with postcode, tenure, floor area in square metres, a sale date and a waterside flag that matches "
        "local geography. In a file of this shape, the business briefing should lead with medians, internal floor area and "
        "grade, and should describe waterfront as a small, high-priced segment. Extreme prices should stay in the analysis "
        "when the question is about that segment.",
    )
    heading(doc, "Digital tools and professional judgement", 2)
    body_paragraph(
        doc,
        "I considered a two-sample t-test because it is the comparison most people reach for first. The normality result "
        "ruled it out. I also considered dropping prices above a Tukey fence so that the second profile would look tidier. "
        "That would have reduced skewness on paper and biased the waterfront comparison, which is the business question. "
        "Leaving the tail in place is the better choice. Seaborn shortened the graphics, and SciPy kept the tests explicit. "
        "A spreadsheet would have hidden the duplicate-id rule and the held-out split. The script is the artefact a "
        "colleague can review, which is also what the appendix is for.",
    )
    body_paragraph(
        doc,
        "The useful next digital step is a geography and dictionary check, not a larger model and not a dashboard. A "
        "dashboard built on a mislabelled file would make the error look more official. Property identifiers can sometimes "
        "be joined back to an address, so this briefing stays at segment level and does not list individual homes. The "
        "result supports a decision about what to trust, what to measure next, and which associations are large enough to "
        "discuss. It does not support a Manchester price forecast. That distinction is the product of the analysis.",
    )

    end = len(doc.paragraphs)
    counted = word_count(doc.paragraphs[start:end])

    heading(doc, "References")
    references = [
        "Harris, C.R., Millman, K.J., van der Walt, S.J. et al. (2020) ‘Array programming with NumPy’, Nature, 585, pp. 357–362.",
        "Hunter, J.D. (2007) ‘Matplotlib: a 2D graphics environment’, Computing in Science & Engineering, 9(3), pp. 90–95.",
        "Kaggle (n.d.) House Sales in King County, USA. Available at: https://www.kaggle.com/datasets/harlfoxem/housesalesprediction (Accessed: 7 November 2026).",
        "Mann, H.B. and Whitney, D.R. (1947) ‘On a test of whether one of two random variables is stochastically larger than the other’, Annals of Mathematical Statistics, 18(1), pp. 50–60.",
        "McKinney, W. (2010) ‘Data structures for statistical computing in Python’, Proceedings of the 9th Python in Science Conference, pp. 56–61.",
        "Pedregosa, F., Varoquaux, G., Gramfort, A. et al. (2011) ‘Scikit-learn: machine learning in Python’, Journal of Machine Learning Research, 12, pp. 2825–2830.",
        "van Rossum, G., Warsaw, B. and Coghlan, N. (2001) PEP 8 – Style guide for Python code. Available at: https://peps.python.org/pep-0008/ (Accessed: 7 November 2026).",
        "Virtanen, P., Gommers, R., Oliphant, T.E. et al. (2020) ‘SciPy 1.0: fundamental algorithms for scientific computing in Python’, Nature Methods, 17, pp. 261–272.",
        "Waskom, M.L. (2021) ‘seaborn: statistical data visualization’, Journal of Open Source Software, 6(60), 3021.",
    ]
    for reference in references:
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.left_indent = Cm(1.25)
        paragraph.paragraph_format.first_line_indent = Cm(-1.25)
        paragraph.paragraph_format.space_after = Pt(6)
        paragraph.paragraph_format.line_spacing = 1.08
        add_text(paragraph, reference, size=11)

    doc.add_page_break()
    heading(doc, "Appendix A. Python code")
    body_paragraph(
        doc,
        "The listing below is the complete script used to produce the audit, the tests and the figures. "
        "It is provided for verification, as required by the assessment brief.",
    )
    source = CODE.read_text(encoding="utf-8").splitlines()
    block = doc.add_paragraph()
    block.paragraph_format.space_after = Pt(0)
    block.paragraph_format.line_spacing = 1.0
    add_text(block, "\n".join(source), name="Consolas", size=8)

    doc.add_page_break()
    heading(doc, "Appendix B. Execution log")
    body_paragraph(doc, "Printed output from the run that generated this report.")
    log_paragraph = doc.add_paragraph()
    log_paragraph.paragraph_format.line_spacing = 1.0
    add_text(log_paragraph, "\n".join(data["log"]), name="Consolas", size=10)

    fill_cover(doc, counted)
    fix_running_header(doc)
    doc.save(OUT)
    print(f"Main-text words: {counted}")
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
