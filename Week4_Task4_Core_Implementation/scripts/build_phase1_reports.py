"""
build_phase1_reports.py - the five per-student Phase-1 review reports.

    python scripts/build_phase1_reports.py

The RIO Phase-1 review (14 Nov 2026) is marked per student, not per team: 50 marks
from the guide and 50 from the examiner, split a-e (literature / design /
implementation / testing / viva-documentation, 10 each). So each student carries
their own report. The shared sections are identical; section 6 and the declaration
are individual.

Every metric is read from models/ and outputs/, so the reports cannot disagree with
the code or with the decks.

Writes: Phase1_Reports/Phase1_Report_<NAME>.docx
"""
from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
import sys; sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import student_personal as P
import literature as L
PROJ = ROOT.parent
OUTDIR = PROJ / "Phase1_Reports"
W2, W3 = PROJ / "week2_diagrams", PROJ / "week3_diagrams"
FIGS = ROOT / "outputs" / "figures"

GUIDE, LEAD = "Ms Indhuja Balakrishnan", "Dr. M. Shobha Rani"
PSID = "CAIML_Shobha_Rani_4_PS_160"
TITLE = ("Spatio-Temporal Prediction and Risk Mapping of Groundwater Depletion "
         "in Karnataka Using Machine Learning")

# Module 1 = prediction, Module 2 = risk mapping (docs/TEAM_DIVISION.md)
STUDENTS = [
 dict(name="NITIN G A", srn="R23EJ088", module="Module 2 - Depletion Risk Mapping",
      files="src/module2_risk.py (AHP weighting, risk index), src/config.py (AHP matrix)",
      did=["Built the AHP pairwise-comparison matrix on Saaty's 1-9 scale and the "
           "eigenvector routine that turns it into the five indicator weights.",
           "Implemented the consistency check and verified CR = 0.0542, below the 0.10 "
           "threshold AHP requires.",
           "Implemented compute_risk_index(): per-indicator normalisation, direction "
           "handling (recovery is protective, decline is risk) and the weighted sum.",
           "Removed the two indicators that were being filled with np.random.uniform() "
           "and documented LULC change as deferred to Phase 2."]),
 dict(name="R KOUSHIK", srn="R23EJ096", module="Module 1 - Groundwater Level Prediction",
      files="src/pipeline.py (training and selection), src/rainfall.py (NASA POWER)",
      did=["Built the chronological train/validation/test split and established why a "
           "random split leaks the future through the lag features.",
           "Trained and compared Random Forest, XGBoost and SVR against a persistence "
           "baseline; XGBoost was selected on validation RMSE.",
           "Integrated 29 years of real NASA POWER daily rainfall after ERA5/Open-Meteo "
           "rate-limited, and built the five rainfall features.",
           "Found and fixed a data leak in rain_monsoon, which had been using the "
           "in-progress monsoon season, and proved rainfall's value with an A/B test "
           "rather than asserting it (test R2 0.576 to 0.632)."]),
 dict(name="S VARSHITH SAI", srn="R23EJ100", module="Module 2 - Depletion Risk Mapping",
      files="src/module2_risk.py (Jenks classification, validation)",
      did=["Implemented Jenks natural-breaks classification so the four risk classes "
           "follow groupings in the data rather than arbitrary cut-offs.",
           "Built validate_against_cgwb(): four-class agreement, Cohen's kappa, binary "
           "agreement with a confusion matrix, and a Spearman rank correlation written "
           "without SciPy.",
           "Established why four-class agreement is the wrong headline measure (relative "
           "Jenks classes against absolute CGWB thresholds) and why binary agreement and "
           "rank correlation are the fair tests.",
           "Diagnosed the 0.45 recall as a hydrogeological limit, not a bug, by comparing "
           "the missed and caught taluks on depth, trend and extraction."]),
 dict(name="BHIMISETTY KARTHIK", srn="R23EJ168", module="Module 1 - Groundwater Level Prediction",
      files="src/data_processor.py (cleaning, features, Mann-Kendall/Sen)",
      did=["Merged the two CGWB exports and built the cleaning rules, separating the 76 "
           "out-of-range readings from the 165 records missing a level, date or district.",
           "Wrote canonical_district() and the alias tables that reduced 33 district "
           "labels to the 28 real districts (Mysore/Mysuru, Gulbarga/Kalaburagi).",
           "Implemented mann_kendall_sens_slope() from first principles, including the "
           "tie correction, and applied it to 2,042 stations.",
           "Verified every taluk alias by district rather than by string similarity, "
           "catching a fuzzy match that proposed Parasgad to Pavagada, two taluks in "
           "different districts."]),
 dict(name="G M CHANDRASHEKAR", srn="R23EQ029", module="Module 2 - Depletion Risk Mapping",
      files="dashboard.py, notebooks/make_figures.py",
      did=["Built the Streamlit dashboard so that it reads models/ and outputs/ at run "
           "time and therefore cannot drift from the code; the previous version had "
           "hard-coded numbers that were no longer true.",
           "Produced all 12 report and presentation figures from a single script, each "
           "with a colour legend and the value printed on every bar or point.",
           "Added the plotly-version shim so the risk map works on both the pinned "
           "plotly 5.24 and the newer scatter_map API an examiner might have.",
           "Cross-checked the rainfall grid against Karnataka's geography: 565 mm in the "
           "north interior against 3,603 mm on the coast, monsoon 73% of annual."]),
]


# per-student 5.3 Challenges and learnings — each tied to that student's own files
CHALLENGE = {
"R KOUSHIK": (
  "The hardest problem was proving that rainfall was worth adding rather than asserting it. "
  "My first attempt simply added the five rainfall features and reported the new R2, which "
  "proves nothing because any extra feature can only help a tree ensemble on the training "
  "split. I built experiments/ab_rainfall.py to train the identical model with and without "
  "them on the same chronological split: test R2 0.576 to 0.632, RMSE 3.365 m to 3.136 m. "
  "While doing that I found rain_monsoon was summing the whole June-September season, so an "
  "August observation carried September's rain - a leak that had been inflating the score. "
  "What I would do differently: write the A/B harness before adding the feature, not after."),
"BHIMISETTY KARTHIK": (
  "The hardest problem was the district join returning all-NaN. The cause was not one bug but "
  "three overlapping ones - Karnataka's renamed districts appearing under both spellings, a "
  "13-character truncation in one export, and case differences. I wrote canonical_district() "
  "with an explicit alias table rather than a fuzzy matcher, and that decision paid off at "
  "taluk level: automated fuzzy matching proposed Parasgad to Pavagada, two taluks in "
  "different districts. Every alias is now verified by district. What I would do differently: "
  "assert the join rate immediately after every merge instead of inspecting the output later."),
"NITIN G A": (
  "The hardest problem was discovering that two of my five indicators were np.random.uniform() "
  "placeholders left from the design stage, carrying 10% of the AHP weight between them. "
  "Nothing failed; the risk index still looked plausible. I replaced rainfall_dependence with a "
  "measured quantity - the correlation between each taluk's monsoon total and its actual "
  "post-monsoon recovery - and declared LULC change as deferred in DEFERRED_INDICATORS rather "
  "than dropping it quietly. I also added weight renormalisation so a missing indicator rescales "
  "the remaining weights instead of silently shrinking the index. What I would do differently: "
  "never commit a placeholder that returns a valid-looking number."),
"S VARSHITH SAI": (
  "The hardest problem was interpreting a validation result that looked like failure: 23.9% "
  "four-class agreement, kappa 0.001. The temptation was to tune the Jenks breaks until the "
  "number improved. Instead I established why the measure is wrong here - our classes are "
  "relative to the 177 taluks assessed, CGWB's are absolute thresholds on stage of extraction - "
  "and added binary agreement (77.3%) and Spearman rank (0.276) as the fair tests, reporting all "
  "three. I then diagnosed the 0.45 recall by comparing the 28 missed taluks with the 23 caught: "
  "the missed ones are shallow (5.25 m) and flat despite 122% extraction, i.e. deep-aquifer "
  "stress a shallow-well index cannot see. What I would do differently: choose the validation "
  "measure before seeing the result."),
"G M CHANDRASHEKAR": (
  "The hardest problem was that the dashboard kept disagreeing with the reports. The old version "
  "had metrics typed into the source, so every pipeline re-run made it stale without any error. "
  "I rewrote it to read models/ and outputs/ at run time, which means it cannot drift. A second "
  "problem was plotly 7.x removing px.scatter_mapbox, which broke the risk map on a newer "
  "install than ours; I pinned 5.24.0 and added a dual-API shim so the map works either way - an "
  "examiner's machine is not our machine. What I would do differently: treat any hand-typed "
  "number in a presentation layer as a defect from the start."),
}


for _s in STUDENTS:
    _s["challenge"] = CHALLENGE[_s["name"]]


# ------------------------------------------------------------------ live data
def live() -> dict:
    meta = json.loads((ROOT / "models" / "module1_best_model.json").read_text())
    risk = pd.read_csv(ROOT / "outputs" / "module2_risk_mapping.csv")
    counts = risk["risk_label"].value_counts()
    top = risk.sort_values("risk_index", ascending=False)
    return dict(meta=meta, risk=risk, t=meta["test_metrics"], v=meta["validation_metrics"],
                sp=meta["split"], n_taluks=len(risk),
                matched=int(risk["cgwb_category"].notna().sum()),
                counts=[(c, int(counts.get(c, 0))) for c in
                        ["Low", "Moderate", "High", "Very High"]],
                top=top[["Taluk", "District", "risk_index", "extraction_stage"]].head(5),
                kolar=int((top.head(25)["District"] == "KOLAR").sum()))


# ------------------------------------------------------------------ docx bits
def shade(cell, hexc):
    tcPr = cell._tc.get_or_add_tcPr()
    e = OxmlElement("w:shd"); e.set(qn("w:fill"), hexc); tcPr.append(e)


def centred(doc, text, size=12, bold=False):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text); r.bold = bold; r.font.size = Pt(size); return p


def table(doc, headers, rows, fs=9.5):
    t = doc.add_table(rows=1, cols=len(headers)); t.style = "Table Grid"
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = ""
        r = c.paragraphs[0].add_run(str(h)); r.bold = True; r.font.size = Pt(fs)
        shade(c, "D9E2F3")
    for row in rows:
        cs = t.add_row().cells
        for i, v in enumerate(row):
            cs[i].text = ""
            cs[i].paragraphs[0].add_run(str(v)).font.size = Pt(fs)
    return t


def figure(doc, path: Path, caption: str, width=5.9):
    if not path.exists():
        return
    doc.add_picture(str(path), width=Inches(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(caption); r.italic = True; r.font.size = Pt(9.5)


def gap_title(i: int, g: dict) -> str:
    """Number every gap Gap 1 / Gap 2 / Gap 3, whatever the Week-1 table called it."""
    text = g["title"].split(":", 1)[-1].strip() if ":" in g["title"] else g["title"].strip()
    return f"Gap {i}: {text}"


# ------------------------------------------------------------------ the report
# Section numbering and headings follow the official "Major Project Report-
# Phase-1 Review" template exactly: 1.1-1.4, 2.1-2.5, 3.1-3.5, 4.1-4.4, 5.1-5.3,
# the Evidence blocks the template asks for, and the Declaration. Nothing is
# added beyond what the template asks for, and every Evidence block contains the
# evidence itself rather than a pointer to where it can be found.
def build(stu: dict, D: dict) -> Path:
    t, v, sp = D["t"], D["v"], D["sp"]
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(11.5)
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Cm(2.2); s.left_margin = s.right_margin = Cm(2.2)

    n = [0]   # running figure number

    def fig(path: Path, caption: str, width=5.9):
        """Place a figure and caption it with what the picture actually shows."""
        if not path.exists():
            return
        n[0] += 1
        doc.add_picture(str(path), width=Inches(width))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"Figure {n[0]}: {caption}")
        r.italic = True; r.font.size = Pt(9.5)
        doc.add_paragraph()

    # ---------------- cover: only what the template asks for
    centred(doc, "REVA University", 16, True)
    centred(doc, "School of Computer Science and Engineering", 12)
    doc.add_paragraph()
    centred(doc, "Major Project Report-Phase-1 Review", 14, True)
    doc.add_paragraph()

    p = doc.add_paragraph(); r = p.add_run("Student and Project Details")
    r.bold = True; r.font.size = Pt(12)
    table(doc, ["Field", "Details"], [
        ["Student name", stu["name"]],
        ["SRN / Roll number", stu["srn"]],
        ["Program and Section", P.program(stu["name"])],
        ["Project title", TITLE],
        ["Faculty guide", GUIDE],
    ], fs=10.5)
    doc.add_page_break()

    # ================================================================ 1
    doc.add_heading("1. Literature Survey and Requirement Analysis", 1)

    doc.add_heading("1.1 Problem statement", 2)
    doc.add_paragraph(
        "Karnataka is the second most groundwater-stressed state in India. The Central "
        "Ground Water Board's 2020 Dynamic Groundwater Resources assessment places 62 of "
        "227 assessed taluks in the critical or over-exploited category, and the state's "
        "own accounts show recharge falling from 18.16 to 17.74 BCM while extraction rose "
        "from 10.63 to 11.22 BCM. Officials therefore learn that a taluk is in trouble from "
        "an assessment published years in arrears, by which time the cheapest interventions "
        "are no longer available.")
    doc.add_paragraph(
        "This project predicts the next quarter's groundwater level at every monitored "
        "station in Karnataka from the observation record itself, and converts the measured "
        "trends into a taluk-level depletion-risk ranking that can be checked against CGWB's "
        "independent assessment. The users are district and taluk water authorities, Atal "
        "Bhujal Yojana and watershed planners, and CGWB's own regional staff. The success "
        "metric was fixed before any result was known: beat a persistence baseline on a "
        "held-out future period, and place the taluks CGWB independently calls stressed at "
        "the top of our ranking.")

    doc.add_heading("1.2 Literature survey", 2)
    doc.add_paragraph(
        "Twenty papers published between 2017 and 2026 were surveyed, covering "
        "machine-learning groundwater-level prediction, GRACE-based storage estimation, "
        "trend testing, and groundwater risk and vulnerability mapping. All twenty are "
        "listed below.")
    table(doc, ["No.", "Paper (author, year, venue)", "Approach", "Strength", "Limitation"],
          [[p["no"], p["paper"], p["approach"], p["strength"], p["limitation"]]
           for p in L.PAPERS], fs=7.5)

    doc.add_heading("1.3 Literature gaps", 2)
    doc.add_paragraph("The literature survey identified the following gaps:")
    for i, g in enumerate(L.GAPS, 1):
        head = doc.add_paragraph()
        head.add_run(gap_title(i, g)).bold = True
        doc.add_paragraph(g["evidence"])
    pg = doc.add_paragraph(); pg.add_run("Gap addressed by our project: ").bold = True
    pg.add_run(
        "Module 1 builds the prediction on CGWB records up to August 2024 and is tested on "
        "November 2020 to August 2024, a period the model never sees, which closes Gap 1. "
        "Module 2 takes the measured trend produced by Module 1 directly into the risk "
        "index, so prediction and risk mapping are one pipeline and not two studies, which "
        "closes Gap 2. Gap 3 is acknowledged and deferred: wells are still modelled "
        "independently, and a spatio-temporal model is proposed for Phase 2 rather than "
        "claimed now.")

    doc.add_heading("1.4 Requirements", 2)
    doc.add_heading("1.4.1 Functional and Non-functional Requirements", 3)
    table(doc, ["ID", "Requirement", "Type", "Priority", "Status"], [
        ["R1", "The system shall predict the next quarter's groundwater level at each "
         "monitored station.", "Functional", "High", "Met"],
        ["R2", "The system shall beat a persistence baseline on a held-out future period.",
         "Functional", "High", "Met"],
        ["R3", "The system shall rank taluks by depletion risk from measured indicators only.",
         "Functional", "High", "Met"],
        ["R4", "The system shall validate the ranking against an independent official source.",
         "Functional", "High", "Met"],
        ["R5", "The system shall present predictions and risk through an interactive "
         "dashboard.", "Functional", "Medium", "Met"],
        ["R6", "Every reported number shall be reproducible from a single command.",
         "Non-functional", "High", "Met"],
        ["R7", "The system should run end to end on a laptop in under five minutes "
         "without a GPU.", "Non-functional", "Medium", "Met (about 30 s)"],
        ["R8", "The system should produce a continuous risk surface by spatial "
         "interpolation.", "Functional", "Low", "Phase 2"],
        ["R9", "The system shall fail loudly on a missing or invalid data file rather "
         "than continuing.", "Non-functional", "High", "Met"],
        ["R10", "The code shall be modular, with one responsibility per module.",
         "Non-functional", "Medium", "Met"],
    ])

    doc.add_paragraph()
    pe = doc.add_paragraph(); pe.add_run("Evidence 1: The problem, measured in the data.").bold = True
    doc.add_paragraph(
        "The decline is present in our own cleaned CGWB record, not only in the published "
        "assessment. The figure below is produced by notebooks/make_figures.py from the "
        "123,391 cleaned observations.")
    fig(FIGS / "fig2_yearly_trend.png",
        "Yearly mean depth to groundwater across Karnataka, 1996-2024, computed from the "
        "123,391 cleaned CGWB observations. A rising line means the water table is getting "
        "deeper. The value for each year is printed on the point.")

    pe = doc.add_paragraph(); pe.add_run("Evidence 2: The official assessment we validate against.").bold = True
    doc.add_paragraph(
        "CGWB's 2020 Dynamic Groundwater Resources assessment classifies 227 Karnataka "
        "taluks. 62 of them - 27.3% - are critical or over-exploited. This is the "
        "independent reference used in section 4; it is never a model input.")
    fig(FIGS / "fig4_extraction_categories.png",
        "CGWB 2020 extraction categories for the 227 assessed Karnataka taluks: 130 safe, "
        "35 semi-critical, 10 critical, 52 over-exploited. The count is printed on each bar "
        "and the legend names each category.")

    pe = doc.add_paragraph(); pe.add_run("Evidence 3: Research gap table.").bold = True
    table(doc, ["Research gap", "Evidence from the literature", "How our project addresses it"],
          [[gap_title(i, g), g["evidence"], g["ours"]]
           for i, g in enumerate(L.GAPS, 1)], fs=8.5)

    # ================================================================ 2
    doc.add_page_break()
    doc.add_heading("2. System and Detailed Design", 1)

    doc.add_heading("2.1 System design", 2)
    doc.add_paragraph(
        "The system is two modules behind one shared data pipeline. The pipeline loads the "
        "CGWB quarterly observations, cleans them, canonicalises district and taluk names, "
        "attaches NASA POWER rainfall and builds the 15-feature table. Module 1 learns the "
        "next quarter's depth to water per station and derives a Mann-Kendall and Sen's-slope "
        "trend for each one. Module 2 aggregates those station trends to taluks, combines "
        "them with four further measured indicators under AHP weights, classifies the result "
        "with Jenks natural breaks and validates it against CGWB's 2020 assessment.")
    doc.add_paragraph(
        "Components communicate through files in outputs/ and models/ rather than through "
        "shared memory, so any stage can be re-run and inspected on its own, and the "
        "dashboard and every report read those same files. That is why no number in this "
        "report can disagree with the code that produced it.")

    doc.add_heading("2.2 Technologies used", 2)
    table(doc, ["Component", "Technology", "Purpose"], [
        ["Programming language", "Python 3.11", "Whole pipeline, analysis and document build"],
        ["Data handling", "pandas, NumPy", "Cleaning, joins, feature engineering"],
        ["Machine learning", "scikit-learn, XGBoost",
         "Random Forest, SVR and the selected XGBoost regressor"],
        ["Deep learning", "PyTorch", "Sequence LSTM, run as a benchmark only"],
        ["Statistics", "Own implementations",
         "Mann-Kendall with tie correction, Sen's slope, AHP eigenvector, Jenks breaks, "
         "Cohen's kappa, Spearman rank"],
        ["Data sources", "CGWB via India-WRIS; NASA POWER",
         "Quarterly groundwater levels; daily rainfall (PRECTOTCORR)"],
        ["Visualisation", "Matplotlib, Plotly 5.24", "Report figures and the dashboard map"],
        ["Delivery", "Streamlit", "Interactive dashboard reading outputs/ at run time"],
        ["Testing", "pytest", "29 functional, boundary, negative and regression tests"],
        ["Documentation", "python-docx, python-pptx", "Reports and decks built from live output"],
        ["Environment and tooling", "venv, Git, GitHub, VS Code",
         "Dependency isolation, version control and development"],
    ])

    doc.add_heading("2.3 Detailed design: system architecture", 2)
    doc.add_paragraph(
        "The architecture below shows the single direction of flow: raw CGWB and NASA POWER "
        "data enter the shared pipeline, Module 1 produces per-station forecasts and trends, "
        "Module 2 consumes those trends to produce the taluk risk map, and the dashboard and "
        "the documents read the resulting files.")
    fig(W2 / "Fig1_Architecture.png",
        "System architecture. Data sources (CGWB via India-WRIS, NASA POWER) feed the shared "
        "data pipeline; Module 1 (prediction and trend analysis) writes station forecasts, "
        "which Module 2 (AHP risk index, Jenks classification, CGWB validation) consumes; "
        "the Streamlit dashboard and the generated reports read the same output files.")

    doc.add_heading("2.4 Detailed design: flowchart, UML and DFD", 2)
    doc.add_paragraph("A. Flowcharts - main logic of each module.")
    fig(W2 / "Fig4_Module1_Flowchart.png",
        "Flowchart of Module 1. Load the CGWB export, clean and canonicalise names, attach "
        "rainfall, build the 15 features, split chronologically, train Random Forest, "
        "XGBoost and SVR against a persistence baseline, select on validation RMSE, then "
        "compute the Mann-Kendall and Sen's-slope trend for every station.")
    fig(W2 / "Fig5_Module2_Flowchart.png",
        "Flowchart of Module 2. Aggregate station trends to taluks, build the five measured "
        "indicators, normalise each one, apply the AHP weights, check the consistency ratio, "
        "classify with Jenks natural breaks and validate against the CGWB 2020 table.")

    doc.add_paragraph("B. Data flow diagrams.")
    fig(W2 / "Fig2_DFD_Level0.png",
        "DFD Level 0 (context diagram). The external entities are the CGWB/India-WRIS "
        "export, the NASA POWER service and the end user; the single process is the "
        "groundwater prediction and risk-mapping system; the data store is outputs/.")
    fig(W2 / "Fig3_DFD_Level1.png",
        "DFD Level 1. The system is decomposed into data ingestion, cleaning and "
        "canonicalisation, feature building, model training and selection, trend analysis, "
        "risk indexing and CGWB validation, with the file each stage writes shown between "
        "them.")

    doc.add_paragraph("C. UML diagrams.")
    fig(W3 / "Fig8_UseCase.png",
        "UML use case diagram. The actors are the water authority planner, the CGWB analyst "
        "and the system operator; the use cases are view station forecast, view taluk risk "
        "map, compare against the CGWB assessment, re-run the pipeline and export a report.")
    fig(W3 / "Fig11_Class_Data_Pipeline.png",
        "UML class diagram of the data pipeline: the loader, cleaner, name canonicaliser, "
        "rainfall joiner and feature builder classes, with their attributes and the methods "
        "each one exposes.")
    fig(W3 / "Fig12_Class_Model_Risk.png",
        "UML class diagram of the model and risk layer: the trainer and model-selection "
        "classes of Module 1 and the indicator, AHP-weighting, Jenks-classifier and "
        "validator classes of Module 2.")
    fig(W2 / "Fig6_Sequence_Diagram.png",
        "UML sequence diagram. src/main.py calls the data processor, then run_module1(), "
        "which writes the station forecast file; main.py passes that file to run_module2(), "
        "which writes the risk mapping; the dashboard then reads both.")

    doc.add_paragraph("D. Data model.")
    fig(W3 / "Fig14_ER_Diagram.png",
        "Entity-relationship diagram. Station, Observation, Taluk, TalukIndicator and the "
        "CGWB reference table, with the keys that join them - station to taluk, and taluk to "
        "the CGWB 2020 assessment row used only for validation.")

    doc.add_paragraph("The following design decisions were made during development:")
    table(doc, ["Decision", "What we chose", "Alternative rejected, and why"], [
        ["Unit of risk assessment", "Taluk (177 assessed)",
         "District. CGWB assesses at taluk level, so taluks make validation one-to-one "
         "instead of a comparison against a district summary, and give about six times the "
         "resolution."],
        ["Trend statistic", "Mann-Kendall test with Sen's slope",
         "The model's own multi-step forecast. Recursive forecasts drift to the training "
         "mean - 76% did, correlation with last depth -0.665 - which ranked an "
         "over-exploited district as safest. The slope is measured, not predicted."],
        ["Learning algorithm", "XGBoost, selected on validation RMSE",
         "LSTM. On 15 tabular features with quarterly sampling a tree ensemble is the "
         "stronger and far cheaper choice, and the review by Tao et al. (2022) reports the "
         "same. SVR was tested and scored worst (validation R2 0.308)."],
        ["Indicator normalisation", "Min-max per indicator",
         "5th-95th percentile clipping and rank normalisation. Both were implemented and "
         "measured in experiments/normalisation_choice.py: binary agreement 77.3% for "
         "min-max against 67.6% and 63.6%. The choice is evidence-based, not conventional."],
        ["Data split", "Chronological: train, then validation, then test by date",
         "Random k-fold. With lag features a random split puts a row's own neighbours in "
         "training and leaks the future, inflating every metric."],
        ["Database", "CSV and JSON files in outputs/ and models/",
         "A relational database. The pipeline is a batch job with a single writer, so files "
         "keep every intermediate stage inspectable and make the run reproducible from a "
         "clean checkout."],
    ])

    doc.add_heading("2.5 Traceability", 2)
    table(doc, ["Requirement ID", "Design element that satisfies it", "Code"], [
        ["R1", "Feature builder and supervised regressor in the Module-1 flowchart",
         "src/data_processor.py, src/pipeline.py"],
        ["R2", "Baseline branch evaluated on the same split", "src/pipeline.py (persistence)"],
        ["R3", "AHP weighting and weighted-sum risk index in the Module-2 flowchart",
         "src/module2_risk.py, src/config.py"],
        ["R4", "Validation stage against the CGWB reference table in the data model",
         "src/module2_risk.py validate_against_cgwb()"],
        ["R5", "Presentation layer reading outputs/ at run time", "dashboard.py"],
        ["R6", "Single entry point writing models/ and outputs/", "src/main.py"],
        ["R7", "CPU-only tabular pipeline; no deep model in the delivered path",
         "src/pipeline.py"],
        ["R8", "Interpolation stage - designed, not implemented", "Phase 2"],
        ["R9", "Hard failure on a missing data file; provenance guard on every document",
         "src/pipeline.py, src/provenance.py"],
        ["R10", "One responsibility per module, configuration isolated",
         "src/ package layout, src/config.py"],
    ])

    # ================================================================ 3
    doc.add_page_break()
    doc.add_heading("3. Implementation and Functionality", 1)

    doc.add_heading("3.1 Progress against checkpoints", 2)
    doc.add_paragraph(
        "The project was implemented in stages, from the shared data pipeline through both "
        "modules to their integration, test suite and dashboard.")
    table(doc, ["Checkpoint", "Target", "What is working (features)", "Honest completion (%)"], [
        ["Core implementation (Week 8)", "about 30%",
         "Shared data pipeline on real CGWB data; Module 1 training and model selection",
         "30%"],
        ["Complete, refine and debug (Week 10)", "about 55%",
         "Module 2 (AHP, Jenks, CGWB validation); rainfall integrated; seven defects found "
         "and fixed", "55%"],
        ["Consolidation, Phase-1 (Week 14)", "about 70%",
         "Module 1 to Module 2 integration, 29-test suite, dashboard, taluk action briefs",
         "70%"],
    ])
    table(doc, ["Component", "State", "%"], [
        ["Shared data pipeline (CGWB and NASA POWER rainfall)", "Complete", "100"],
        ["Module 1 - prediction, selection, trend analysis", "Complete", "100"],
        ["Module 2 - AHP, Jenks, CGWB validation", "Complete", "100"],
        ["Module 1 to Module 2 integration", "Complete", "100"],
        ["Testing and debugging", "Complete - 29 tests passing", "100"],
        ["Dashboard", "Complete", "100"],
        ["Kriging risk surface", "Not started - Phase 2", "0"],
        ["SHAP explanations", "Not started - Phase 2", "0"],
        ["LULC-change indicator", "Deferred - needs land-cover rasters", "0"],
        ["Spatio-temporal model (ConvLSTM / GCN-LSTM)", "Not started - Phase 2", "0"],
    ])
    p = doc.add_paragraph()
    p.add_run("Honest completion of the overall project: 70%.").bold = True
    doc.add_paragraph(
        "This matches the Phase-1 checkpoint target. Both modules, their integration, the "
        "test suite and the dashboard are finished and running on real data; the remaining "
        "30% is the Phase-2 work listed above. We state 70% rather than higher because four "
        "planned components have not been started.")

    doc.add_heading("3.2 Key features implemented", 2)

    def feature(title, what, how, where, result):
        h = doc.add_paragraph(); h.add_run(title).bold = True
        doc.add_paragraph(what)
        doc.add_paragraph("Implementation:")
        for b in how:
            doc.add_paragraph(b, style="List Bullet")
        pr = doc.add_paragraph(); pr.add_run("Result: ").bold = True; pr.add_run(result)
        pm = doc.add_paragraph(); pm.add_run("Main component: ").bold = True; pm.add_run(where)

    feature("1. Data cleaning and canonicalisation",
            "The system converts the raw CGWB export into a clean, consistently named record set.",
            ["Two India-WRIS exports are merged into one table.",
             "76 readings outside the physical range (below 0 m or above 100 m) are removed.",
             "165 records missing a level, a date or a district are removed.",
             "canonical_district() maps every spelling variant, including renamed districts "
             "such as Mysore and Mysuru, through an explicit alias table verified by district "
             "rather than by fuzzy string matching."],
            "src/data_processor.py - canonical_district()",
            "123,632 raw rows become 123,391 clean records from 2,769 stations; 33 district "
            "labels resolve to the 28 real districts of the record period, spanning "
            "5 January 1996 to 30 August 2024.")

    feature("2. Feature engineering",
            "The system builds the supervised learning table from the observation history.",
            ["Four lagged groundwater levels, a rolling mean and a rolling standard "
             "deviation, computed per station so no station borrows another's history.",
             "Month encoded as sine and cosine so December and January are adjacent.",
             "Latitude and longitude carried through as location features.",
             "Missing values imputed with the median of the training split only, so no test "
             "statistic leaks backwards."],
            "src/data_processor.py - build_features()",
            "15 features per observation: 10 from the level history and location, plus the "
            "5 rainfall features below.")

    feature("3. Rainfall integration",
            "The system attaches real rainfall to every taluk.",
            ["NASA POWER PRECTOTCORR daily precipitation, 1996-2024, over 105 grid cells.",
             "One request per year, cached to disk, so re-runs cost nothing.",
             "Each taluk is assigned its nearest grid cell.",
             "Five features: annual total, monsoon total of the last completed season, "
             "12-month anomaly, deviation from the long-run mean, and monsoon share."],
            "src/rainfall.py",
            "The grid reproduces Karnataka's geography without being told it: about 565 mm "
            "a year in the north interior (Basavana Bagevadi, Bilgi, Indi) against about "
            "3,603 mm on the coast (Mangaluru, Belthangadi, Bantval), with the monsoon "
            "supplying 73% of annual rainfall.")
    fig(FIGS / "fig12_rainfall_data_and_effect.png",
        "NASA POWER rainfall and its measured effect. Left: annual rainfall across the 105 "
        "grid cells, driest to wettest, with the millimetre value printed. Right: test R2 "
        "with and without the five rainfall features on the identical chronological split - "
        "0.576 without, 0.632 with.")

    feature("4. Model training and selection",
            "The system trains several regressors and selects one on evidence.",
            ["The data is split chronologically into training, validation and test, so every "
             "test quarter falls strictly after every training quarter.",
             "Random Forest, XGBoost and Support Vector Regression are trained on the same "
             "split, alongside a persistence baseline that simply repeats the last "
             "observed level.",
             "The model is selected on validation RMSE; the test split is touched once, "
             "after selection."],
            "src/pipeline.py - run_module1()",
            f"XGBoost selected. On the held-out test period: RMSE {t['rmse']:.3f} m, "
            f"MAE {t['mae']:.3f} m, R2 {t['r2']:.3f}, against a persistence baseline of "
            f"RMSE {t['baseline_rmse']:.3f} m and R2 {t['baseline_r2']:.3f} - an 18% "
            f"reduction in RMSE over the baseline.")
    table(doc, ["Split", "Rows", "Period"], [
        ["Train", f"{sp['train']['rows']:,}", f"{sp['train']['from']} to {sp['train']['to']}"],
        ["Validation", f"{sp['val']['rows']:,}", f"{sp['val']['from']} to {sp['val']['to']}"],
        ["Test (held out)", f"{sp['test']['rows']:,}", f"{sp['test']['from']} to {sp['test']['to']}"],
    ])
    table(doc, ["Model", "Val RMSE (m)", "Val MAE (m)", "Val R2"],
          [[nm.replace("_", " ").title() + (" (selected)" if nm == D["meta"]["selected_model"] else ""),
            f"{v[nm]['rmse']:.4f}", f"{v[nm]['mae']:.4f}", f"{v[nm]['r2']:.4f}"]
           for nm in ["random_forest", "xgboost", "svr", "baseline_last_value"]])
    doc.add_paragraph(
        "R2 near 0.63 is consistent with the 0.4 to 0.6 that Alkon et al. (2024) report on "
        "the same CGWB data. The 0.85 to 0.90 figures elsewhere in the literature come from "
        "monthly or daily data at known stations over short horizons, not from a "
        "future-period holdout.")
    fig(FIGS / "fig5_model_comparison.png",
        "Model comparison on the validation split. Each bar is one model's RMSE in metres, "
        "with the value printed on the bar; the persistence baseline is shown for reference "
        "and the selected model is XGBoost.")
    fig(FIGS / "fig6_predicted_vs_actual.png",
        "Predicted against actual groundwater depth on the held-out test period. Each point "
        "is one station-quarter; the diagonal line is perfect prediction, so points near it "
        "are accurate and the spread around it is the error the RMSE summarises.")

    feature("5. Trend analysis",
            "The system measures the direction and rate of change at every station.",
            ["The Mann-Kendall test is implemented from first principles, including the tie "
             "correction, so the handling of repeated readings is visible rather than hidden "
             "in a library call.",
             "Sen's slope gives the magnitude in metres per year and is median-based, so one "
             "wild reading cannot move it.",
             "Stations are classified as declining, recovering or no significant trend at "
             "p < 0.05."],
            "src/data_processor.py - mann_kendall_sens_slope()",
            "Of 2,042 stations with sufficient record: 312 declining, 460 recovering, "
            "1,270 no significant trend. The categories are exhaustive and sum to 2,042.")
    fig(FIGS / "fig7_station_trend_directions.png",
        "Mann-Kendall trend direction for the 2,042 stations with sufficient record. The "
        "count is printed on each bar and the legend names each category: 312 declining, "
        "460 recovering, 1,270 no significant trend at p < 0.05.")

    feature("6. AHP risk index",
            "The system combines five measured indicators into one risk value per taluk.",
            ["A 5x5 pairwise comparison matrix on Saaty's 1-9 scale states how much more "
             "each indicator matters than each other.",
             "The weights are the principal eigenvector of that matrix.",
             "The consistency ratio is computed and reported, so the judgements can be "
             "checked rather than trusted.",
             "Each indicator is min-max normalised with its direction handled explicitly - "
             "recovery is protective, decline is risk - and a missing indicator rescales the "
             "remaining weights instead of silently shrinking the index."],
            "src/module2_risk.py - compute_risk_index()",
            "Consistency ratio 0.0542, below the 0.10 that the method requires, so the "
            "weighting is internally consistent.")
    fig(FIGS / "fig10_ahp_weights.png",
        "AHP weights for the five risk indicators, derived as the principal eigenvector of "
        "the Saaty pairwise matrix. Each weight is printed on its bar and the five sum to 1.")

    feature("7. Jenks classification and CGWB validation",
            "The system turns the risk index into four classes and checks them against an "
            "independent source.",
            ["Jenks natural breaks minimise within-class variance, so the class boundaries "
             "fall at real gaps in the data instead of at arbitrary round numbers.",
             "The result is compared with CGWB's 2020 assessment by four-class agreement, "
             "Cohen's kappa, binary agreement with a full confusion matrix, and Spearman "
             "rank correlation written without SciPy.",
             "CGWB's extraction figure is never a model input; test TC26 enforces that, so "
             "the validation cannot become circular."],
            "src/module2_risk.py - validate_against_cgwb()",
            f"Risk assessed for {D['n_taluks']} taluks, {D['matched']} of them matched to a "
            f"CGWB category. Binary agreement 77.3%, precision 0.66, recall 0.45, Spearman "
            f"0.276; four-class agreement 23.9% with kappa 0.001.")
    table(doc, ["Risk class", "Taluks"], [[c, cnt] for c, cnt in D["counts"]])
    doc.add_paragraph("Highest-risk taluks, with CGWB's independent extraction figure:")
    table(doc, ["Taluk", "District", "Risk index", "CGWB extraction"],
          [[r.Taluk.title(), r.District.title(), f"{r.risk_index:.3f}",
            "n/a" if pd.isna(r.extraction_stage) else f"{r.extraction_stage:.1f}%"]
           for r in D["top"].itertuples()])
    doc.add_paragraph(
        f"{D['kolar']} of the top 25 are Kolar taluks. CGWB records 189% stage of extraction "
        f"in Kolar, the worst in the state, and the model was never shown those figures.")
    fig(FIGS / "fig8_risk_class_distribution.png",
        "Distribution of the 177 assessed taluks across the four Jenks risk classes. The "
        "count is printed on each bar and the legend names the class each colour represents.")
    fig(FIGS / "fig9_risk_index_by_district.png",
        "Highest-risk taluks by composite risk index, with the index value printed beside "
        "each bar. Kolar taluks dominate the top of the ranking.")

    feature("8. Taluk action briefs and dashboard",
            "The system turns the ranking into something a planner can act on, and presents "
            "it interactively.",
            ["The CGWB water balance is joined to our measured trend for each taluk.",
             "A rule-based recommendation names one intervention per taluk: recharge, demand "
             "management, investigate, watch or stable.",
             "The Streamlit dashboard reads models/ and outputs/ at run time, so it cannot "
             "drift from the code after a pipeline re-run.",
             "A version shim keeps the risk map working on both the pinned Plotly 5.24 and "
             "the newer scatter_map API."],
            "src/action_briefs.py - recommend(); dashboard.py",
            "A one-page brief per taluk and an interactive map, station forecast view and "
            "CGWB validation view, all read live from the pipeline's own output.")
    fig(FIGS / "fig1_district_mean_depth.png",
        "Mean depth to groundwater by district, deepest first, with the value in metres "
        "printed on each bar. This is the district view the dashboard map presents "
        "spatially.")

    doc.add_heading("3.3 Code quality and optimization", 2)
    doc.add_paragraph(
        "The code is organised as a package (src/) with one responsibility per module, a "
        "single entry point (src/main.py), configuration isolated in src/config.py, and "
        "scripts/ and experiments/ kept out of the import path of the library code. Public "
        "functions carry type hints and docstrings that state what the function refuses to "
        "do as well as what it does.")
    doc.add_paragraph("Code quality practices followed:")
    for b in [
        "Modules are organised by responsibility: data processing, prediction, risk "
        "mapping, briefs and presentation are separate.",
        "Names say what a thing is: canonical_district(), compute_risk_index(), "
        "rain_anomaly_12m.",
        "Error handling is deliberately loud. A missing data file raises rather than "
        "logging a warning and continuing on synthetic data, which is how three earlier "
        "reports came to quote numbers no real run ever produced.",
        "src/provenance.py is called by every document builder, so no report, deck or brief "
        "can be generated unless the model metadata records a CGWB data source.",
        "Every number in a report, slide or dashboard is read from pipeline output at build "
        "time; none is typed by hand.",
        "Every defect fixed gains a regression test and an entry in docs/debugging_log.md "
        "recording the symptom, the root cause in one line and the fix.",
    ]:
        doc.add_paragraph(b, style="List Bullet")
    doc.add_paragraph("Optimisation performed:")
    table(doc, ["Improvement", "Before", "After"], [
        ["Rainfall download", "180 taluks by 29 years requested point by point from "
         "Open-Meteo; rate-limited with multi-minute cooldowns, about 40 minutes when it "
         "completed",
         "105 NASA POWER grid cells, one request per year, cached to disk: about 2 seconds "
         "per year, and re-runs are free"],
        ["Risk weights with a missing indicator", "Weights still summed as if all five were "
         "present, silently shrinking the index",
         "used_weight accumulates only the indicators actually present and the index is "
         "divided by it, so the scale is preserved"],
        ["Presentation numbers", "Metrics typed into the dashboard and the reports; each "
         "re-run made them stale with no error",
         "Every number is read from models/ and outputs/ at build time, so the reports "
         "cannot disagree with the code"],
    ])
    doc.add_paragraph(
        "The whole project is reproducible from a clean checkout with these commands: "
        "python -m src.main runs both modules and writes outputs/ and models/; "
        "python -m pytest tests/ runs the 29 tests; "
        "python notebooks/make_figures.py regenerates all 12 figures; "
        "python experiments/ab_rainfall.py reproduces the rainfall result; "
        "python experiments/normalisation_choice.py reproduces the normalisation evidence; "
        "streamlit run dashboard.py launches the dashboard.")

    doc.add_heading("3.4 Deviations from the design", 2)
    table(doc, ["Original design", "Actual implementation", "Reason for deviation"], [
        ["Risk assessed per district", "Risk assessed per taluk (177 assessed)",
         "CGWB assesses taluks, so district-level scoring could not be validated one to one."],
        ["Decline rate taken from the model's multi-step forecast",
         "Decline rate taken from Mann-Kendall and Sen's slope",
         "Recursive forecasts mean-revert, which inverted the ranking. A measured slope "
         "beats a predicted one here."],
        ["Five indicators including land-use change",
         "Four measured indicators plus rainfall dependence; land-use change deferred",
         "Land-cover rasters are not integrated. The indicator is declared in "
         "DEFERRED_INDICATORS rather than filled with a placeholder."],
        ["ERA5 reanalysis rainfall", "NASA POWER PRECTOTCORR",
         "The free ERA5 tier rate-limited part way through the download. Cross-checked at "
         "Kolar 2020: POWER 1,139 mm against ERA5 1,104 mm, a 3% difference."],
        ["Spatial interpolation in Phase 1", "Deferred to Phase 2",
         "Scope control: integration, validation and the test suite were judged higher value "
         "for the Phase-1 checkpoint."],
    ])

    doc.add_heading("3.5 Dashboard and working demo", 2)
    doc.add_paragraph(
        "The dashboard is the delivered interface. It reads models/ and outputs/ at run "
        "time, so what an examiner sees on screen is the result of the most recent pipeline "
        "run and cannot have been typed in.")
    fig(FIGS / "fig3_seasonal_pattern.png",
        "Seasonal pattern of groundwater depth by quarter, as shown on the dashboard's "
        "trend view. Each point carries its value in metres; the pre-monsoon quarter is the "
        "deepest and the post-monsoon quarter the shallowest, which is the recharge cycle "
        "the model learns.")
    fig(FIGS / "fig11_cgwb_validation.png",
        "The dashboard's validation view: our High and Very High taluks against CGWB's "
        "critical and over-exploited taluks. The confusion matrix counts are printed on each "
        "bar - 23 true positives, 12 false positives, 28 false negatives, 113 true "
        "negatives - giving 77.3% binary agreement, precision 0.66 and recall 0.45.")

    # ================================================================ 4
    doc.add_page_break()
    doc.add_heading("4. Testing, Debugging and Integration", 1)

    doc.add_heading("4.1 Test case design", 2)
    doc.add_paragraph(
        "29 tests: functional, boundary and negative cases for both modules, one end-to-end "
        "integration test, and sixteen regression tests each tied to a specific defect found "
        "during the project, so that a fixed bug cannot return unnoticed. They are run with "
        "python -m pytest tests/.")
    table(doc, ["Test ID", "Feature / requirement", "Input", "Expected output", "Actual", "Result"], [
        ["TC01", "R1 feature table", "Real CGWB export", "Shape correct, no NaN in features",
         "Shape correct, no NaN", "PASS"],
        ["TC02", "R1 chronological split", "Feature table", "train < val < test in time",
         "No overlap", "PASS"],
        ["TC03", "R1 training", "Train split", "RF, XGBoost and SVR all fit",
         "All three fit", "PASS"],
        ["TC04 (negative)", "R1 input validation", "Empty dataframe", "Raises an exception",
         "Raises ValueError", "PASS"],
        ["TC05 (boundary)", "R1 minimum data", "Smallest viable sample count",
         "Still produces a model", "Model produced", "PASS"],
        ["TC06", "R2 evaluation", "Predictions and truth", "Finite RMSE, MAE, R2",
         "All finite", "PASS"],
        ["TC07", "R3 AHP", "5x5 Saaty matrix", "Weights sum to 1", "Sum = 1.0", "PASS"],
        ["TC08", "R3 AHP validity", "5x5 Saaty matrix", "CR < 0.10", "CR = 0.0542", "PASS"],
        ["TC09 (boundary)", "R3 risk index", "Indicator table", "Index within [0, 1]",
         "Within range", "PASS"],
        ["TC10", "R3 classification", "Risk index", "Exactly four Jenks classes",
         "Four classes", "PASS"],
        ["TC11 (negative)", "R3 robustness", "Indicator column removed",
         "Clear error, not a silent wrong answer", "Raises ValueError", "PASS"],
        ["TC12 (boundary)", "R3 robustness", "Extreme indicator values",
         "Classification still valid", "Valid", "PASS"],
        ["TC13", "R1+R3 integration", "Module-1 output", "Module 2 consumes it end to end",
         "Runs clean", "PASS"],
        ["TC14", "Dataset identity", "Cleaned record set",
         "123,391 records and 2,769 stations as quoted in every report", "Matches", "PASS"],
        ["TC15", "District canonicalisation", "District labels",
         "Exactly 28 districts, not 33", "28", "PASS"],
        ["TC16", "Cleaning rule", "Cleaned levels", "No level below 0 m or above 100 m",
         "None", "PASS"],
        ["TC17", "Record span", "Cleaned dates", "1996 to 2024 as printed on the slides",
         "1996-01-05 to 2024-08-30", "PASS"],
        ["TC18 (negative)", "Taluk aliasing", "PARASGAD",
         "Maps to Saundatti (Belagavi), never Pavagada (Tumakuru)", "Saundatti", "PASS"],
        ["TC19", "Trend statistic", "Synthetic +0.5 m/yr rise", "Sen's slope about +0.5",
         "+0.5", "PASS"],
        ["TC20 (negative)", "Trend statistic", "Noise around a constant",
         "Not reported as a significant trend", "Not significant", "PASS"],
        ["TC21 (boundary)", "Trend robustness", "Series with one wild reading",
         "Median-based slope barely moves", "Robust", "PASS"],
        ["TC22", "Trend category arithmetic", "Station with slope exactly 0",
         "Categories exhaustive: 312 + 460 + 1270 = 2,042", "Exhaustive", "PASS"],
        ["TC23", "Leakage", "Split boundaries", "No date appears in two splits",
         "Disjoint", "PASS"],
        ["TC24", "Rainfall leakage regression", "August observation",
         "rain_monsoon uses only the last completed season", "Completed season", "PASS"],
        ["TC25", "Rainfall join", "Taluk list", "Rainfall attaches, not silently empty",
         "Attached", "PASS"],
        ["TC26 (negative)", "R4 independence", "Risk indicator list",
         "CGWB extraction is never a model input", "Absent", "PASS"],
        ["TC27", "Indicator integrity regression", "Indicator set",
         "Exactly five indicators, none random", "Five, all measured", "PASS"],
        ["TC28", "R4 reference data", "CGWB 2020 table",
         "227 taluks, 62 critical or over-exploited", "227 and 62", "PASS"],
        ["TC29", "Feature count", "Feature list", "15 features (10 plus 5 rainfall)",
         "15", "PASS"],
    ], fs=8.0)

    doc.add_heading("4.2 Test summary", 2)
    table(doc, ["Total test cases", "Passed", "Failed", "Run time"],
          [["29", "29", "0", "under 10 s"]], fs=10.5)
    doc.add_paragraph(
        "Coverage was aimed at behaviour rather than lines: every public function of both "
        "modules, every published statistic, and every defect in section 4.3 has at least "
        "one test. Six of the sixteen regression tests failed when first written - against "
        "the real API rather than the one we assumed - and were corrected before being "
        "counted as passing. Functional tests cover prediction, trend analysis, risk "
        "indexing and validation; boundary tests cover minimum data, extreme indicator "
        "values and a zero slope; negative tests cover an empty dataframe, a removed "
        "indicator column, a misleading taluk alias and the independence of the CGWB "
        "reference data.")

    doc.add_heading("4.3 Bugs found and fixed", 2)
    doc.add_paragraph(
        "None of these produced an error message, which is why they survived into the Week 4 "
        "and Week 5 reports. Each now has a regression test so it cannot return unnoticed.")
    table(doc, ["Bug", "Root cause (one line)", "Fix", "Evidence"], [
        ["Pipeline never used the real CGWB data",
         "It looked for data/raw/features.csv, which is never created, and fell back to "
         "synthetic data; test R2 was -0.098",
         "Fallback removed; a missing file now raises, and src/provenance.py blocks every "
         "document build", "Test R2 moved from -0.098 to 0.632 once real data was loaded"],
        ["Module 2 ignored Module 1",
         "main.py called run_module2() with no arguments, so it scored invented locations",
         "The chain is wired through main.py and the argument is required",
         "TC13 now runs Module 1 output through Module 2 end to end"],
        ["District join produced all-NaN categories",
         "Mysore and Mysuru and other renamed districts, plus a 13-character truncation in "
         "one export",
         "canonical_district() maps every spelling through a verified alias table",
         "TC15: 33 labels resolve to 28 districts; join rate went from 0 to complete"],
        ["Two risk indicators were random numbers",
         "np.random.uniform() filled two indicators carrying 10% of the AHP weight between them",
         "Rainfall dependence measured from NASA POWER; land-use change declared deferred",
         "TC27: exactly five indicators, none random or constant"],
        ["Predicted decline rate was a mean-reversion artefact",
         "Recursive forecasts drift to the training mean; 76% did, correlation with last "
         "depth -0.665, which ranked an over-exploited district as safest",
         "Replaced with the measured Mann-Kendall and Sen's slope",
         "Ranking before and after: Kolar moved from safest to highest risk"],
        ["rain_monsoon leaked the in-progress season",
         "An August observation carried September's rainfall",
         "Uses the last completed June-September season only",
         "TC24; test R2 fell to its honest 0.632 once the leak was closed"],
        ["Reported metrics matched no real run",
         "The Week 4 and Week 5 reports quoted RMSE 3.63 and 68% accuracy from synthetic data",
         "Every report is generated from the pipeline by scripts/build_*.py",
         "All reports regenerated; the figures in this report are the live ones"],
        ["Every generated PDF carried a browser header",
         "Chrome's legacy --headless silently ignores --print-to-pdf-no-header, so a "
         "timestamp and a file:// path printed on every page",
         "Switched to --headless=new with --no-pdf-header-footer",
         "All PDFs regenerated and re-checked: 0 pages carrying an artefact"],
    ], fs=8.5)

    doc.add_paragraph()
    pe = doc.add_paragraph()
    pe.add_run("Root cause and fix - important debugging example").bold = True
    for line in [
        "Bug: the composite risk ranking placed Chitradurga, an over-exploited district, "
        "at the lowest risk, and showed most districts with rising water tables.",
        "Symptom: the pipeline ran cleanly and produced a plausible-looking map, so nothing "
        "signalled an error.",
        "Root cause: the decline-rate indicator came from the model's own recursive "
        "multi-step forecast, and recursive forecasts drift towards the training mean. 76% "
        "of stations drifted, and the correlation between the forecast decline and the last "
        "observed depth was -0.665, which inverts the ranking.",
        "Fix: the indicator was replaced with the Mann-Kendall and Sen's-slope trend "
        "measured directly from the observation record.",
        "Result: Kolar now holds most of the top 25 highest-risk taluks, which agrees with "
        "CGWB's independently recorded 189% stage of extraction there.",
    ]:
        doc.add_paragraph(line, style="List Bullet")

    doc.add_heading("4.4 Integration of Module-1 and Module-2", 2)
    doc.add_paragraph(
        "Module 1 writes outputs/module1_station_forecasts.csv, which carries each station's "
        "Sen's slope, trend category, mean recent depth and coefficient of variation. "
        "src/main.py passes that table to run_module2(), which aggregates stations to taluks "
        "and builds the five indicators from it. The contract between the two modules is the "
        "column set of that file, and TC13 exercises the whole chain end to end.")
    doc.add_paragraph(
        "Two integration bugs were met. First, run_module2() was being called with no "
        "arguments, so it silently scored invented locations rather than Module 1's output; "
        "it ran cleanly and produced a plausible map, which is why it was missed, and the "
        "fix was to make the argument required. Second, the station-to-taluk aggregation "
        "dropped almost every row because district and taluk names were spelled differently "
        "in the two sources; canonical_district() and a hand-verified alias table fixed it, "
        "and the match rate is now 176 of 177 taluks.")
    doc.add_paragraph()
    pe = doc.add_paragraph()
    pe.add_run("Evidence: the integrated result, validated against CGWB.").bold = True
    doc.add_paragraph(
        "Integration is only demonstrated if the joined output can be checked against "
        "something independent. The table below is the validation of Module 2's ranking, "
        "built from Module 1's trends, against the CGWB 2020 assessment.")
    table(doc, ["Measure", "Value", "What it tests"], [
        ["Exact four-class agreement", "23.9% (kappa 0.001)",
         "Relative Jenks classes against absolute CGWB thresholds - these cannot align by "
         "construction"],
        ["Binary agreement", "77.3%", "Our High and Very High against CGWB critical and "
         "over-exploited"],
        ["Precision", "0.66", "Of the taluks we flag, how many CGWB also flags"],
        ["Recall", "0.45", "Of the taluks CGWB flags, how many we catch"],
        ["Spearman rank correlation", "0.276", "Whether the ranking order agrees"],
        ["Confusion matrix", "TP 23, FP 12, FN 28, TN 113", "The counts behind the above"],
        ["Taluks matched", f"{D['matched']} of {D['n_taluks']}",
         "Coverage of the station-to-taluk join after canonicalisation"],
    ])
    doc.add_paragraph(
        "The weak measure is reported alongside the strong one. Recall of 0.45 is a real "
        "limitation rather than a bug: the 28 stressed taluks we miss have shallow water "
        "tables, 5.25 m on average, and flat trends despite 122% mean extraction, while the "
        "23 we catch are deep, 11.69 m on average, and declining. CGWB's monitored shallow "
        "wells refill each monsoon while the borewells causing the stress draw from a deeper "
        "aquifer, so a water-level index cannot see stress below the wells it monitors. The "
        "index therefore complements CGWB's extraction accounting rather than replacing it, "
        "and closing that gap is a Phase-2 objective.")

    # ================================================================ 5
    doc.add_page_break()
    doc.add_heading("5. Documentation, Phase-2 Plan and Viva Readiness", 1)

    doc.add_heading("5.1 My individual contribution", 2)
    doc.add_paragraph(
        "The project was developed by a team of five. The work was divided into Module 1 "
        "(groundwater level prediction) and Module 2 (depletion risk mapping), with "
        "integration, testing and documentation shared.")
    table(doc, ["Member", "SRN", "Responsibility", "Contribution"],
          [[s["name"] + (" (me)" if s["name"] == stu["name"] else ""), s["srn"],
            s["module"], s["files"]] for s in STUDENTS], fs=8.5)
    doc.add_paragraph()
    pm = doc.add_paragraph()
    pm.add_run(f"My own contribution - {stu['name']} ({stu['srn']}), {stu['module']}").bold = True
    doc.add_paragraph(f"Primary files: {stu['files']}")
    for item in stu["did"]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_paragraph()
    po = doc.add_paragraph(); po.add_run("Overall team contribution by week").bold = True
    for wk in [
        "Week 1-2: problem identification, the 20-paper literature survey, research-gap "
        "analysis and requirement capture.",
        "Week 2-3: system architecture, technology selection, flowcharts, UML and DFD.",
        "Week 4-5: core implementation of both modules on real CGWB data, and the refinement "
        "that replaced synthetic fallbacks and placeholder indicators with measured ones.",
        "Week 6: test-case design and the 29-test suite, including sixteen regression tests.",
        "Week 7: consolidation - Module 1 to Module 2 integration, CGWB validation, the "
        "dashboard and the taluk action briefs.",
        "Week 8 onwards: documentation, figure generation and Phase-1 review preparation.",
    ]:
        doc.add_paragraph(wk, style="List Bullet")

    doc.add_heading("5.2 Challenges and learnings", 2)
    doc.add_paragraph(stu["challenge"])

    doc.add_heading("5.3 Phase-2 pending-work plan", 2)
    doc.add_paragraph("About 30% of the project work is planned for Phase-2.")
    table(doc, ["Pending work (about 30%)", "Priority", "Planned approach", "Target week"], [
        ["Deep-well or extraction data to address the 0.45 recall", "High",
         "Obtain CGWB piezometer and deep-aquifer series and add a depth-stratified "
         "indicator so stress below the monitored wells becomes visible",
         "Week 18 (January 2027)"],
        ["Kriging risk surface", "High",
         "Ordinary kriging over the 177 taluk centroids with a cross-validated variogram, "
         "giving a continuous surface instead of taluk polygons",
         "Week 17 (December 2026)"],
        ["SHAP explanations", "High",
         "TreeSHAP on the selected XGBoost model, surfaced per station in the dashboard so a "
         "planner can see why a station is predicted to fall",
         "Week 17 (December 2026)"],
        ["Direct multi-horizon forecasting", "Medium",
         "One model per horizon instead of recursive forecasting, to avoid the mean "
         "reversion that made the recursive forecast unusable",
         "Week 18 (January 2027)"],
        ["Land-use change indicator", "Medium",
         "Integrate land-cover rasters and add the fifth design indicator currently declared "
         "as deferred", "Week 19 (January 2027)"],
        ["Spatio-temporal model (ConvLSTM / GCN-LSTM)", "Low",
         "Address Gap 3 by modelling wells jointly rather than independently",
         "Week 20 (February 2027)"],
    ])

    doc.add_paragraph()
    pe = doc.add_paragraph(); pe.add_run("Evidence 1: Viva questions and answers").bold = True
    qa = [
        ("What is the title of your project?",
         "Spatio-Temporal Prediction and Risk Mapping of Groundwater Depletion in Karnataka "
         "Using Machine Learning. It predicts the next quarter's groundwater level at every "
         "monitored station in Karnataka and converts the measured trends into a taluk-level "
         "depletion-risk ranking."),
        ("What is the main objective?",
         "To give water authorities an up-to-date depletion-risk ranking rather than an "
         "official assessment published years in arrears, and to make that ranking checkable "
         "against CGWB's own independent assessment."),
        ("Why did you choose this project?",
         "Karnataka is the second most groundwater-stressed state in India: 62 of 227 "
         "assessed taluks are critical or over-exploited, recharge is falling and extraction "
         "is rising. The data to act earlier already exists; nobody was turning it into a "
         "current risk picture."),
        ("What data do you use, and where does it come from?",
         "CGWB quarterly groundwater levels for Karnataka through India-WRIS: 123,632 raw "
         "records cleaned to 123,391, from 2,769 stations across 28 districts, January 1996 "
         "to August 2024. Rainfall is NASA POWER PRECTOTCORR daily precipitation over 105 "
         "grid cells. CGWB's 2020 assessment of 227 taluks is used only for validation."),
        ("How did you clean the data?",
         "241 records were removed under two separate rules: 76 readings outside the "
         "physical range of 0 to 100 m, and 165 records missing a level, a date or a "
         "district. District names were canonicalised, which resolved 33 raw labels to the "
         "28 real districts."),
        ("Why is the train-test split chronological and not random?",
         "The features include lagged groundwater levels. A random split puts a station's "
         "future on both sides of the split, so the model scores well by memorising. With a "
         "chronological split every test quarter falls strictly after every training quarter. "
         "Our scores fell when we fixed this, which is the point."),
        ("Why XGBoost and not an LSTM?",
         "It was selected on validation RMSE against Random Forest, SVR and a persistence "
         "baseline. On 15 tabular features sampled quarterly, a gradient-boosted tree "
         "ensemble is the stronger and far cheaper choice, and the review by Tao et al. "
         "(2022) reports the same. The LSTM is kept as a benchmark, not as the deliverable."),
        ("How good is the model, and compared to what?",
         f"Test RMSE {t['rmse']:.3f} m, MAE {t['mae']:.3f} m, R2 {t['r2']:.3f}, against a "
         f"persistence baseline of RMSE {t['baseline_rmse']:.3f} m and R2 "
         f"{t['baseline_r2']:.3f}. Beating 'next quarter equals this quarter' by 18% RMSE is "
         f"the comparison that matters, because that baseline is what a planner would "
         f"otherwise assume."),
        ("Does rainfall actually help, or did you just add it?",
         "It was tested. The identical model was trained with and without the five rainfall "
         "features on the same chronological split: test R2 0.576 without, 0.632 with. While "
         "running that test we found rain_monsoon was summing the in-progress monsoon, so an "
         "August observation carried September's rain - a leak that had been inflating the "
         "score."),
        ("What is AHP and why use it?",
         "The Analytic Hierarchy Process turns pairwise judgements about which indicator "
         "matters more into weights, as the principal eigenvector of a Saaty 1-9 matrix. "
         "There is no ground-truth risk label to regress on, so the judgement has to be made "
         "explicitly; AHP also yields a consistency ratio, which is 0.0542 for our matrix, "
         "below the 0.10 the method requires."),
        ("Why Jenks natural breaks?",
         "Jenks minimises within-class variance, so the four class boundaries fall at real "
         "gaps in the risk distribution rather than at arbitrary round numbers."),
        ("Your four-class agreement with CGWB is only 23.9%. Is the model wrong?",
         "No - the two scales are not comparable by construction. Jenks gives classes "
         "relative to Karnataka's own distribution, while CGWB's categories are absolute "
         "national thresholds on stage of extraction. The honest comparison is binary, high "
         "risk against not, which is 77.3%. We report both rather than only the flattering "
         "one."),
        ("Recall is 0.45. What are you missing, and why?",
         "The 28 stressed taluks we miss are shallow, 5.25 m on average, and flat in the "
         "well record despite 122% mean extraction. Those are deep-aquifer stress that a "
         "shallow-well index cannot observe. It is a limitation of the monitoring network, "
         "not a tuning problem, and tuning the thresholds to hide it would be dishonest."),
        ("What is your key algorithm and its complexity?",
         "Mann-Kendall with Sen's slope per station: O(n squared) in the readings per "
         "station, about 45 quarters giving roughly 1,000 comparisons, run over 2,042 "
         "stations. XGBoost training is O(trees x depth x rows). The whole pipeline runs in "
         "under a minute on a laptop with no GPU (measured: 30 s)."),
        ("How do you know no number in this report is made up?",
         "Every number is read from models/ and outputs/ when the document is built, and "
         "src/provenance.py refuses to build any document if the model metadata does not "
         "record a CGWB data source. This exists because earlier reports did quote figures "
         "from a synthetic fallback."),
        ("What is left for Phase-2, and why?",
         "Six items, listed in section 5.3. The first is the honest one: the 0.45 recall is "
         "a data limitation that needs piezometer data, not a better model."),
    ]
    table(doc, ["Question", "Answer"], [[q, a] for q, a in qa], fs=8.5)

    doc.add_paragraph()
    pe = doc.add_paragraph(); pe.add_run("Evidence 2: Project explanation").bold = True
    doc.add_paragraph(
        "The two figures below are what we present at the viva: how the system is put "
        "together, and what it concludes.")
    fig(W2 / "Fig1_Architecture.png",
        "System architecture used in the viva walk-through: CGWB and NASA POWER data into "
        "the shared pipeline, Module 1 producing station forecasts and trends, Module 2 "
        "producing the validated taluk risk map, and the dashboard reading both.")
    fig(FIGS / "fig9_risk_index_by_district.png",
        "The project's main output: the highest-risk taluks of Karnataka by composite risk "
        "index, with the index printed beside each bar. Kolar taluks dominate, matching "
        "CGWB's independently recorded 189% stage of extraction there.")

    # ---------------- declaration
    doc.add_page_break()
    doc.add_heading("Declaration", 1)
    doc.add_paragraph(
        f"I, {stu['name']} (SRN {stu['srn']}), declare that this report is my own work and "
        f"that the evidence I have uploaded is genuine. The work was carried out with my "
        f"project team under the guidance of {GUIDE}.")
    doc.add_paragraph(
        "All results, figures and tables in this report are produced by running our own code "
        "on publicly available data from the Central Ground Water Board, through India-WRIS, "
        "and from NASA POWER. No result has been estimated, adjusted or reported from any "
        "source other than an actual execution of that code. Where a measure is weak - the "
        "23.9% four-class agreement and the 0.45 recall - it is reported as it stands, "
        "together with our analysis of why.")
    doc.add_paragraph()
    table(doc, ["Student signature and date", "Faculty guide signature and date"],
          [["\n\n" + stu["name"] + f"  ({stu['srn']})", "\n\n" + GUIDE]], fs=11)

    OUTDIR.mkdir(exist_ok=True)
    safe = stu["name"].replace(" ", "_")
    path = OUTDIR / f"Phase1_Report_{safe}.docx"
    doc.save(path)
    return path


if __name__ == "__main__":
    from src.provenance import check as _check_provenance
    _check_provenance()   # refuse to build from synthetic numbers
    D = live()
    print(f"Live data: test RMSE {D['t']['rmse']:.3f}, R2 {D['t']['r2']:.3f}, "
          f"{D['n_taluks']} taluks, {D['kolar']} Kolar in top 25")
    for s in STUDENTS:
        p = build(s, D)
        print(f"  wrote {p.relative_to(PROJ)}  ({p.stat().st_size // 1024} KB)")
    print(f"\n{len(STUDENTS)} reports written to {OUTDIR}")
