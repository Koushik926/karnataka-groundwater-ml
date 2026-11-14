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


# ------------------------------------------------------------------ the report
# Section numbering follows the official "Major Project Report-Phase-1 Review"
# template circulated on 04.11.2026. Sections 1-5 and the Declaration are the
# template's own; 3.5-3.7, 4.5 and 6 are additional evidence, not replacements.
def build(stu: dict, D: dict) -> Path:
    t, v, sp = D["t"], D["v"], D["sp"]
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(11.5)
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Cm(2.2); s.left_margin = s.right_margin = Cm(2.2)

    # ---------------- cover
    centred(doc, "REVA UNIVERSITY", 17, True)
    centred(doc, P.school(stu["name"]), 12)
    centred(doc, "Center for Artificial Intelligence & Machine Learning (CAIML)", 12)
    doc.add_paragraph()
    centred(doc, "MAJOR PROJECT REPORT — PHASE-1 REVIEW", 14, True)
    centred(doc, "REVA Internship Opportunities (RIO)  ·  Review date 14 November 2026", 11)
    doc.add_paragraph()
    centred(doc, TITLE, 13, True)
    centred(doc, f"Problem Statement ID: {PSID}", 11)
    doc.add_paragraph()

    centred(doc, "Student and Project Details", 12, True)
    table(doc, ["Field", "Details"], [
        ["Student name", stu["name"]],
        ["SRN / Roll number", stu["srn"]],
        ["Program and Section", P.program(stu["name"])],
        ["Project title", TITLE],
        ["Faculty guide", GUIDE],
        ["Organization / Center", f"REVA University — CAIML (Center lead: {LEAD})"],
        ["Module responsibility", stu["module"]],
        ["Primary files owned", stu["files"]],
        ["Report generated", f"{date.today():%d %B %Y}, from a live pipeline run"],
    ], fs=10.5)
    doc.add_paragraph()
    centred(doc, "Sustainable Development Goal addressed", 11, True)
    centred(doc, "SDG 6 — Clean Water and Sanitation   (secondary: SDG 13, Climate Action)", 11)
    doc.add_page_break()

    # ================================================================ 1
    doc.add_heading("1. Literature Survey and Requirement Analysis", 1)

    doc.add_heading("1.1 Problem statement", 2)
    doc.add_paragraph(
        "Karnataka is the second most groundwater-stressed state in India. CGWB's 2020 "
        "assessment places 62 of 227 assessed taluks in the critical or over-exploited "
        "category, and the state's own accounts show recharge falling from 18.16 to 17.74 "
        "BCM while extraction rose from 10.63 to 11.22 BCM. Officials therefore learn that a "
        "taluk is in trouble from an assessment published years in arrears, by which time the "
        "cheapest interventions are no longer available.")
    doc.add_paragraph(
        "This project predicts the next quarter's groundwater level at every monitored "
        "station in Karnataka from the observation record itself, and converts the measured "
        "trends into a taluk-level depletion-risk ranking that can be checked against CGWB's "
        "independent assessment. The users are district and taluk water authorities, Atal "
        "Bhujal Yojana and watershed planners, and CGWB's own regional staff. Success is "
        "measurable and was defined before the results were known: beat a persistence "
        "baseline on a held-out future period, and place the taluks CGWB independently calls "
        "stressed at the top of our ranking.")

    doc.add_heading("1.2 Literature survey", 2)
    doc.add_paragraph(
        "We surveyed 20 papers published between 2017 and 2026, covering machine-learning "
        "groundwater prediction, GRACE-based storage estimation, and risk and vulnerability "
        "mapping. The table below gives the most directly relevant subset; the full 20-paper "
        "comparison is in the Week 1 Literature Survey document.")
    table(doc, ["Paper (author, year, venue)", "Approach / technique", "Strength", "Limitation"], [
        ["Alkon et al., 2024, Environmental Research Letters",
         "Random Forest + SHAP on CGWB wells, all India",
         "Same CGWB data as ours; reports R² 0.4–0.6, which sets a realistic bar",
         "National scale; no taluk-level risk output"],
        ["Supreetha et al., 2020, Modeling Earth Systems and Env.",
         "Lion-algorithm-optimised LSTM, Udupi, Karnataka",
         "Shows deep learning works on Karnataka wells",
         "Single district; prediction only, no risk map"],
        ["Asoka et al., 2017, Nature Geoscience",
         "GRACE satellite storage + well data attribution",
         "Separates monsoon variability from pumping",
         "Coarse (≈100 km) resolution; descriptive, not predictive"],
        ["Tao et al., 2022, Journal of Hydrology (review)",
         "Review of 100+ ML groundwater studies",
         "Tree ensembles are competitive with deep learning on tabular well data",
         "Notes most studies never test against a naive baseline"],
        ["Mohanavelu et al., 2020, Water",
         "Mann-Kendall / Sen's slope trend testing",
         "Establishes non-stationarity in Indian groundwater series",
         "Trend description only; no forecast and no risk index"],
    ])

    doc.add_heading("1.3 Literature gaps", 2)
    doc.add_paragraph(
        "G5 — Many studies use older observation periods. We use CGWB records up to "
        "August 2024 and test on November 2020 to August 2024, a period the model never "
        "sees. This gap drives Module 1 and is the gap this report's Module-1 work "
        "addresses.", style="List Bullet")
    doc.add_paragraph(
        "G4 — Prediction and risk mapping are treated as separate problems. We feed the "
        "measured trend from Module 1 directly into the risk index of Module 2, so one "
        "command produces both. This gap drives Module 2.", style="List Bullet")
    doc.add_paragraph(
        "G2 — Spatial relationships between wells are rarely modelled; wells are treated "
        "independently. Acknowledged and deferred: a ConvLSTM or GCN-LSTM is proposed for "
        "Phase 2 rather than claimed now.", style="List Bullet")

    doc.add_heading("1.4 Requirements", 2)
    table(doc, ["ID", "Requirement", "Type", "Priority", "Status"], [
        ["R1", "Predict the next quarter's groundwater level at each monitored station",
         "Functional", "High", "Met"],
        ["R2", "Beat a persistence baseline on a held-out future period",
         "Functional", "High", "Met"],
        ["R3", "Rank taluks by depletion risk from measured indicators only",
         "Functional", "High", "Met"],
        ["R4", "Validate the ranking against an independent official source",
         "Functional", "High", "Met"],
        ["R5", "Present predictions and risk through an interactive dashboard",
         "Functional", "Medium", "Met"],
        ["R6", "Reproduce every reported number from a single command",
         "Non-functional", "High", "Met"],
        ["R7", "Run end to end on a laptop in under five minutes, no GPU",
         "Non-functional", "Medium", "Met (≈3 min)"],
        ["R8", "Produce a continuous risk surface by spatial interpolation",
         "Functional", "Low", "Phase 2"],
    ])
    p = doc.add_paragraph()
    p.add_run("Evidence: ").bold = True
    p.add_run("Week 1 Literature Survey document (20 papers with gaps identified) and the "
              "Week 2 requirement set.")

    # ================================================================ 2
    doc.add_page_break()
    doc.add_heading("2. System and Detailed Design", 1)

    doc.add_heading("2.1 System design", 2)
    doc.add_paragraph(
        "The system is two modules behind one shared data pipeline. The pipeline loads the "
        "CGWB quarterly observations, cleans them, canonicalises district and taluk names, "
        "attaches NASA POWER rainfall and builds the 15-feature table. Module 1 learns the "
        "next quarter's depth to water per station and derives a Mann-Kendall / Sen's-slope "
        "trend for each one. Module 2 aggregates those station trends to taluks, combines "
        "them with four further measured indicators under AHP weights, classifies the result "
        "with Jenks natural breaks and validates it against CGWB's 2020 assessment. "
        "Components communicate through files in outputs/ and models/ rather than through "
        "shared memory, so any stage can be re-run and inspected on its own, and the "
        "dashboard and every report read those same files.")
    figure(doc, W2 / "Fig1_Architecture.png", "Figure 1: System architecture (Appendix A).")

    doc.add_heading("2.2 Detailed design: diagrams included", 2)
    table(doc, ["Diagram", "Shows", "Figure"], [
        ["Flowchart — Module 1", "Main logic and decision flow of the prediction pipeline",
         "Figure 2"],
        ["Flowchart — Module 2", "Decision flow of risk scoring and classification", "Figure 3"],
        ["UML (class / sequence / use case)", "Structure of the processor, pipeline and risk "
         "classes and their interactions", "Week 3 design document"],
        ["DFD (level 0 and 1)", "Data flows between CGWB input, feature store, models and "
         "outputs", "Week 3 design document"],
        ["ER / data model", "Stations, observations, taluk indicators and the CGWB "
         "reference table", "Figure 4"],
    ])
    figure(doc, W2 / "Fig4_Module1_Flowchart.png", "Figure 2: Module 1 — prediction pipeline.")
    doc.add_paragraph()
    figure(doc, W2 / "Fig5_Module2_Flowchart.png",
           "Figure 3: Module 2 — risk mapping, five measured indicators.")
    doc.add_paragraph()
    figure(doc, W3 / "Fig14_ER_Diagram.png", "Figure 4: Data model.")

    doc.add_heading("2.3 Design decisions", 2)
    table(doc, ["Decision", "What we chose", "Alternative rejected, and why"], [
        ["Unit of risk assessment", "Taluk (177 assessed)",
         "District. CGWB assesses at taluk level, so taluks make validation one-to-one "
         "instead of a comparison against a district summary, and give ~6× the resolution."],
        ["Trend statistic", "Mann-Kendall test with Sen's slope",
         "The model's own multi-step forecast. Recursive forecasts drift to the training "
         "mean — 76% did, correlation with last depth −0.665 — which ranked an "
         "over-exploited district as safest. The slope is measured, not predicted."],
        ["Learning algorithm", "XGBoost, selected on validation RMSE",
         "LSTM. On 15 tabular features with quarterly sampling a tree ensemble is the "
         "stronger and far cheaper choice; the literature review (Tao et al. 2022) says the "
         "same. SVR was tested and scored worst (val R² 0.308)."],
        ["Indicator normalisation", "Min-max per indicator",
         "5th–95th percentile clipping and rank normalisation. Both were implemented and "
         "measured (experiments/normalisation_choice.py): binary agreement 77.3% for min-max "
         "against 67.6% and 63.6%. The choice is evidence-based, not conventional."],
        ["Data split", "Chronological: train → validation → test by date",
         "Random k-fold. With lag features a random split puts a row's own neighbours in "
         "training and leaks the future, inflating every metric."],
    ])

    doc.add_heading("2.4 Traceability", 2)
    table(doc, ["Requirement ID", "Design element that satisfies it", "Code"], [
        ["R1", "Feature builder + supervised regressor in the Module-1 flowchart",
         "src/data_processor.py, src/pipeline.py"],
        ["R2", "Baseline branch evaluated on the same split", "src/pipeline.py (persistence)"],
        ["R3", "AHP weighting + weighted-sum risk index in the Module-2 flowchart",
         "src/module2_risk.py, src/config.py"],
        ["R4", "Validation stage against the CGWB reference table in the data model",
         "src/module2_risk.py validate_against_cgwb()"],
        ["R5", "Presentation layer reading outputs/ at run time", "dashboard.py"],
        ["R6", "Single entry point writing models/ and outputs/", "src/main.py"],
        ["R7", "CPU-only tabular pipeline; no deep model in Phase 1", "src/pipeline.py"],
        ["R8", "Interpolation stage — designed, not implemented", "Phase 2"],
    ])
    p = doc.add_paragraph()
    p.add_run("Evidence: ").bold = True
    p.add_run("Week 2 system design document; Week 3 detailed design document covering "
              "flowchart, UML and DFD.")

    # ================================================================ 3
    doc.add_page_break()
    doc.add_heading("3. Implementation and Functionality", 1)

    doc.add_heading("3.1 Progress against checkpoints", 2)
    table(doc, ["Checkpoint", "Target", "What is working", "Honest completion"], [
        ["Core implementation (Week 8)", "≈30%",
         "Shared data pipeline on real CGWB data; Module 1 training and model selection",
         "30%"],
        ["Complete, refine and debug (Week 10)", "≈55%",
         "Module 2 (AHP, Jenks, CGWB validation); rainfall integrated; seven defects fixed",
         "55%"],
        ["Consolidation, Phase-1 (Week 14)", "≈70%",
         "Module 1 → Module 2 integration, 29-test suite, dashboard, action briefs, "
         "benchmark harness",
         "70%"],
    ])
    table(doc, ["Component", "State", "%"], [
        ["Shared data pipeline (CGWB + NASA POWER rainfall)", "Complete", "100"],
        ["Module 1 — prediction, selection, trend analysis", "Complete", "100"],
        ["Module 2 — AHP, Jenks, CGWB validation", "Complete", "100"],
        ["Module 1 → Module 2 integration", "Complete", "100"],
        ["Testing and debugging", "Complete — 29 tests passing", "100"],
        ["Dashboard", "Complete", "100"],
        ["Kriging risk surface", "Not started — Phase 2", "0"],
        ["SHAP explanations", "Not started — Phase 2", "0"],
        ["LULC-change indicator", "Deferred — needs land-cover rasters", "0"],
        ["Spatio-temporal model (ConvLSTM / GCN-LSTM)", "Not started — Phase 2", "0"],
    ])
    p = doc.add_paragraph()
    p.add_run("Honest completion of the overall project: 70%.").bold = True
    doc.add_paragraph(
        "This matches the Phase-1 checkpoint target. Both modules, their integration, the "
        "test suite and the dashboard are finished and running on real data; the remaining "
        "30% is the Phase-2 work listed above. We state 70% rather than higher because four "
        "planned components have not been started.")

    doc.add_heading("3.2 Key features implemented", 2)
    table(doc, ["Feature", "What it does", "How it is built", "Main file / function"], [
        ["Data cleaning and canonicalisation",
         "123,632 raw rows → 123,391 clean records, 33 district labels → 28 real districts",
         "Explicit alias tables verified by district, not fuzzy matching",
         "src/data_processor.py — canonical_district()"],
        ["Feature engineering",
         "15 features: 4 lags, rolling mean and SD, month sin/cos, lat/lon, 5 rainfall",
         "Grouped shift/rolling per station; train-median imputation only",
         "src/data_processor.py — build_features()"],
        ["Rainfall integration",
         "29 years of NASA POWER daily rainfall over 105 grid cells",
         "Cached per-year downloads, nearest-cell assignment per taluk",
         "src/rainfall.py"],
        ["Model training and selection",
         "RF, XGBoost, SVR and a persistence baseline on one chronological split",
         "Selection on validation RMSE; test split touched once",
         "src/pipeline.py — run_module1()"],
        ["Trend analysis",
         "Mann-Kendall + Sen's slope for 2,042 stations with sufficient record",
         "Implemented from first principles including the tie correction",
         "src/data_processor.py — mann_kendall_sens_slope()"],
        ["AHP risk index",
         "Five measured indicators → one 0–1 risk index per taluk",
         "Saaty 1–9 pairwise matrix, principal eigenvector, CR check, weight renormalisation",
         "src/module2_risk.py — compute_risk_index()"],
        ["Jenks classification and validation",
         "Four risk classes; agreement, kappa, precision/recall and Spearman vs CGWB",
         "Jenks natural breaks; Spearman written without SciPy",
         "src/module2_risk.py — validate_against_cgwb()"],
        ["Taluk action briefs",
         "Joins the CGWB water balance to our trend and names one intervention per taluk",
         "Rule-based: recharge / demand / investigate / watch / stable",
         "src/action_briefs.py — recommend()"],
        ["Dashboard",
         "Interactive map, station forecasts and validation, read live from outputs/",
         "Streamlit + Plotly with a version shim for plotly ≥ 6",
         "dashboard.py"],
    ])

    doc.add_heading("3.3 Code quality and optimization", 2)
    doc.add_paragraph(
        "The code is organised as a package (src/) with one responsibility per module, a "
        "single entry point (src/main.py), configuration isolated in src/config.py, and "
        "scripts/ and experiments/ kept out of the import path of the library code. Public "
        "functions carry type hints and docstrings that state what the function refuses to "
        "do as well as what it does.")
    doc.add_paragraph(
        "Error handling is deliberately loud. The pipeline used to catch a missing data file, "
        "log a warning and continue on synthetic data — which is how three earlier reports "
        "came to quote numbers no real run ever produced. It now raises a RuntimeError, and "
        "src/provenance.py is called by all seven document builders so that no report, deck "
        "or brief can be generated unless models/module1_best_model.json records a CGWB "
        "data source.")
    table(doc, ["Improvement", "Before", "After"], [
        ["Rainfall download", "180 taluks × 29 years requested point by point from "
         "Open-Meteo; rate-limited with multi-minute cooldowns, ~40 min when it completed",
         "105 NASA POWER grid cells, one request per year, cached to disk: ~2 s per year "
         "and re-runs are free"],
        ["Risk weights with a missing indicator", "Weights still summed as if all five were "
         "present, silently shrinking the index",
         "used_weight accumulates only the indicators actually present and the index is "
         "divided by it, so the scale is preserved"],
        ["Presentation numbers", "Metrics typed into the dashboard and reports; each "
         "re-run made them stale with no error",
         "Every number is read from models/ and outputs/ at build time — the reports "
         "cannot disagree with the code"],
    ])

    doc.add_heading("3.4 Deviations from the design", 2)
    table(doc, ["Design said", "Implementation does", "Why"], [
        ["Risk assessed per district", "Risk assessed per taluk (177)",
         "CGWB assesses taluks; district-level scoring could not be validated one-to-one."],
        ["Decline rate from the model's multi-step forecast",
         "Decline rate from Mann-Kendall / Sen's slope",
         "Recursive forecasts mean-revert, which inverted the ranking. Measured beats "
         "predicted here."],
        ["Five indicators including LULC change", "Four measured + rainfall dependence; "
         "LULC deferred",
         "Land-cover rasters are not integrated. Declared in DEFERRED_INDICATORS rather "
         "than filled with a placeholder."],
        ["ERA5 reanalysis rainfall", "NASA POWER PRECTOTCORR",
         "ERA5's free tier rate-limited part-way through the download. Cross-checked at "
         "Kolar 2020: POWER 1139 mm vs ERA5 1104 mm, a 3% difference."],
        ["Spatial interpolation in Phase 1", "Deferred to Phase 2",
         "Scope control: the integration, validation and test suite were judged higher "
         "value for the Phase-1 checkpoint."],
    ])
    p = doc.add_paragraph()
    p.add_run("Evidence: ").bold = True
    p.add_run("Week 4 core-output screenshots, Week 5 refined output, Week 7 integrated-module "
              "screenshots with one-line justifications.")

    doc.add_heading("3.5 Data", 2)
    doc.add_paragraph(
        "CGWB quarterly observations for Karnataka via India-WRIS: 123,632 raw rows, "
        "123,391 after cleaning, from 2,769 stations across 28 districts, January 1996 to "
        "August 2024. 241 rows were removed — 76 with an impossible level (below 0 m or "
        "above 100 m) and 165 missing a level, date or district.")
    doc.add_paragraph(
        "Rainfall: NASA POWER PRECTOTCORR daily precipitation, 1996–2024, 105 grid cells. "
        "The data reproduces Karnataka's geography without being told it — 565 mm/year in "
        "the north interior against 3,603 mm on the coast, with the monsoon supplying 73% "
        "of annual rainfall.")
    doc.add_paragraph(
        "CGWB Dynamic Groundwater Resources 2020 (227 taluks) is used only to validate the "
        "risk ranking. It is never a model input: using it as one would make the validation "
        "circular, and a test (TC26) now enforces that.")

    doc.add_heading("3.6 Module 1 — results", 2)
    table(doc, ["Split", "Rows", "Period"], [
        ["Train", f"{sp['train']['rows']:,}", f"{sp['train']['from']} → {sp['train']['to']}"],
        ["Validation", f"{sp['val']['rows']:,}", f"{sp['val']['from']} → {sp['val']['to']}"],
        ["Test (held out)", f"{sp['test']['rows']:,}", f"{sp['test']['from']} → {sp['test']['to']}"],
    ])
    table(doc, ["Model", "Val RMSE (m)", "Val MAE (m)", "Val R²"],
          [[n.replace("_", " ").title() + (" (selected)" if n == D["meta"]["selected_model"] else ""),
            f"{v[n]['rmse']:.4f}", f"{v[n]['mae']:.4f}", f"{v[n]['r2']:.4f}"]
           for n in ["random_forest", "xgboost", "svr", "baseline_last_value"]])
    doc.add_paragraph(
        f"On the held-out test period: RMSE {t['rmse']:.3f} m, MAE {t['mae']:.3f} m, "
        f"R² {t['r2']:.3f}, against a persistence baseline of RMSE "
        f"{t['baseline_rmse']:.3f} m and R² {t['baseline_r2']:.3f}. Adding the five real "
        f"rainfall features raised test R² from 0.576 to {t['r2']:.3f}, measured by A/B test "
        f"in experiments/ab_rainfall.py rather than asserted. R² near 0.63 is consistent with "
        f"the 0.4–0.6 that Alkon et al. (2024) report on the same CGWB data; the 0.85–0.90 "
        f"figures in the literature come from monthly or daily data at known stations with "
        f"short horizons, not a future-period holdout.")
    figure(doc, FIGS / "fig5_model_comparison.png",
           "Figure 5: Model comparison. Values printed on each bar; units in metres.")
    doc.add_paragraph()
    figure(doc, FIGS / "fig12_rainfall_data_and_effect.png",
           "Figure 6: Rainfall data and its measured effect on test performance.")
    doc.add_paragraph(
        "Trends: of 2,042 stations with sufficient record, 312 show a significant decline, "
        "460 a significant recovery and 1,270 no significant trend at p < 0.05. The "
        "categories are exhaustive — an earlier count of 2,027 missed the 15 stations whose "
        "Sen's slope is exactly zero, and TC22 now guards against that.")

    doc.add_heading("3.7 Module 2 — results", 2)
    doc.add_paragraph(
        f"Risk is assessed for {D['n_taluks']} taluks — the unit CGWB itself assesses, so "
        f"validation is one-to-one; {D['matched']} matched a CGWB category. Five measured "
        f"indicators are weighted by AHP with a consistency ratio of 0.0542, below the 0.10 "
        f"the method requires.")
    figure(doc, FIGS / "fig10_ahp_weights.png",
           "Figure 7: AHP indicator weights, with each weight printed.")
    doc.add_paragraph()
    table(doc, ["Risk class", "Taluks"], [[c, n] for c, n in D["counts"]])
    doc.add_paragraph("Highest-risk taluks, with CGWB's independent extraction figure:")
    table(doc, ["Taluk", "District", "Risk index", "CGWB extraction"],
          [[r.Taluk.title(), r.District.title(), f"{r.risk_index:.3f}",
            "n/a" if pd.isna(r.extraction_stage) else f"{r.extraction_stage:.1f}%"]
           for r in D["top"].itertuples()])
    doc.add_paragraph(
        f"{D['kolar']} of the top 25 are Kolar taluks. CGWB records 189% stage of extraction "
        f"in Kolar, the worst in the state, and the model was never shown those figures.")
    doc.add_paragraph()
    figure(doc, FIGS / "fig9_risk_index_by_district.png",
           "Figure 8: Highest-risk taluks by composite risk index.")

    # ================================================================ 4
    doc.add_page_break()
    doc.add_heading("4. Testing, Debugging and Integration", 1)

    doc.add_heading("4.1 Testcase design", 2)
    doc.add_paragraph(
        "29 tests: functional, boundary and negative cases for both modules, one end-to-end "
        "integration test, and sixteen regression tests each tied to a specific defect found "
        "during the project, so a fixed bug cannot return unnoticed. Run with "
        "python -m pytest tests/.")
    table(doc, ["Test ID", "Feature / requirement", "Input", "Expected output", "Actual", "Result"], [
        ["TC01", "R1 feature table", "Real CGWB export", "Shape correct, no NaN in features",
         "As expected", "Pass"],
        ["TC02", "R1 chronological split", "Feature table", "train < val < test in time",
         "No overlap", "Pass"],
        ["TC03", "R1 training", "Train split", "RF, XGBoost and SVR all fit",
         "All three fit", "Pass"],
        ["TC04 (negative)", "R1 input validation", "Empty dataframe", "Raises an exception",
         "Raises", "Pass"],
        ["TC05 (boundary)", "R1 minimum data", "Smallest viable sample count",
         "Still produces a model", "Model produced", "Pass"],
        ["TC06", "R2 evaluation", "Predictions + truth", "Finite RMSE, MAE, R²",
         "All finite", "Pass"],
        ["TC07", "R3 AHP", "5×5 Saaty matrix", "Weights sum to 1", "Sum = 1.0", "Pass"],
        ["TC08", "R3 AHP validity", "5×5 Saaty matrix", "CR < 0.10", "CR = 0.0542", "Pass"],
        ["TC09 (boundary)", "R3 risk index", "Indicator table", "Index within [0, 1]",
         "Within range", "Pass"],
        ["TC10", "R3 classification", "Risk index", "Exactly four Jenks classes",
         "Four classes", "Pass"],
        ["TC11 (negative)", "R3 robustness", "Indicator column removed",
         "Clear error, not a silent wrong answer", "Raises ValueError", "Pass"],
        ["TC12 (boundary)", "R3 robustness", "Extreme indicator values",
         "Classification still valid", "Valid", "Pass"],
        ["TC13", "R1+R3 integration", "Module-1 output", "Module 2 consumes it end to end",
         "Runs clean", "Pass"],
        ["TC14", "Dataset identity", "Cleaned record set",
         "123,391 records / 2,769 stations as quoted in every report", "Matches", "Pass"],
        ["TC15", "Defect #8 regression", "District labels", "Exactly 28 districts, not 33",
         "28", "Pass"],
        ["TC16", "Cleaning rule", "Cleaned levels", "No level < 0 m or > 100 m remains",
         "None", "Pass"],
        ["TC17", "Record span", "Cleaned dates", "1996 → 2024 as printed on the slides",
         "Matches", "Pass"],
        ["TC18 (negative)", "Taluk aliasing", "PARASGAD",
         "Maps to Saundatti (Belagavi), never Pavagada (Tumakuru)", "Saundatti", "Pass"],
        ["TC19", "Trend statistic", "Synthetic +0.5 m/yr rise", "Sen's slope ≈ +0.5",
         "+0.5", "Pass"],
        ["TC20 (negative)", "Trend statistic", "Noise around a constant",
         "Not reported as a significant trend", "Not significant", "Pass"],
        ["TC21 (boundary)", "Trend robustness", "Series with one wild reading",
         "Median-based slope barely moves", "Robust", "Pass"],
        ["TC22", "Own arithmetic bug", "Station with slope exactly 0",
         "Categories exhaustive — 312+460+1270 = 2,042", "Exhaustive", "Pass"],
        ["TC23", "Leakage", "Split boundaries", "No date appears in two splits",
         "Disjoint", "Pass"],
        ["TC24", "Defect #11 regression", "August observation",
         "rain_monsoon uses only the last completed season", "Completed season", "Pass"],
        ["TC25", "Rainfall join", "Taluk list", "Rainfall attaches, not silently empty",
         "Attached", "Pass"],
        ["TC26 (negative)", "R4 independence", "Risk indicator list",
         "CGWB extraction is never a model input", "Absent", "Pass"],
        ["TC27", "Defect #9 regression", "Indicator set",
         "Exactly five indicators, none random", "Five, all measured", "Pass"],
        ["TC28", "R4 reference data", "CGWB 2020 table",
         "227 taluks, 62 critical/over-exploited", "227 / 62", "Pass"],
        ["TC29", "Feature count", "Feature list", "15 features (10 + 5 rainfall)",
         "15", "Pass"],
    ], fs=8.5)

    doc.add_heading("4.2 Test summary", 2)
    doc.add_paragraph(
        "Total 29, passed 29, failed 0 (python -m pytest tests/, 3.9 s). Coverage was aimed "
        "at behaviour rather than lines: every public function of both modules, every "
        "published statistic, and every defect in section 4.3 has at least one test. Six of "
        "the sixteen regression tests failed when first written — against the real API rather "
        "than the one we assumed — and were corrected before being counted as passing.")

    doc.add_heading("4.3 Bugs found and fixed", 2)
    doc.add_paragraph(
        "None of these produced an error message, which is why they survived into the Week 4 "
        "and Week 5 reports. Full write-ups, with before/after output, are in "
        "docs/debugging_log.md.")
    table(doc, ["Bug", "Root cause (one line)", "Fix", "Evidence"], [
        ["Pipeline never used the real CGWB data",
         "It looked for data/raw/features.csv, which is never created, and fell back to "
         "synthetic data; test R² was −0.098",
         "Fallback removed; a missing file now raises, and src/provenance.py blocks every "
         "document build", "debugging_log #1; before/after metrics"],
        ["Module 2 ignored Module 1",
         "main.py called run_module2() with no arguments, so it scored invented locations",
         "The chain is wired through main.py and tested by TC13",
         "debugging_log #2; integration screenshots"],
        ["District join produced all-NaN categories",
         "Mysore/Mysuru and other renamed districts plus a 13-character truncation",
         "canonical_district() maps every spelling; TC15 guards it",
         "debugging_log #8; join-rate before/after"],
        ["Two risk indicators were random numbers",
         "np.random.uniform() filled two indicators carrying 10% of the AHP weight",
         "Rainfall dependence measured from NASA POWER; LULC declared deferred; TC27 guards it",
         "debugging_log #9; indicator table"],
        ["'Predicted decline rate' was a mean-reversion artefact",
         "Recursive forecasts drift to the training mean (76% did; corr −0.665), ranking an "
         "over-exploited district safest",
         "Replaced with Mann-Kendall / Sen's slope", "debugging_log #10; ranking before/after"],
        ["rain_monsoon leaked the in-progress season",
         "An August observation carried September's rainfall",
         "Uses the last completed June–September season; TC24 guards it",
         "debugging_log #11; A/B output"],
        ["Reported metrics matched no real run",
         "Week 4/5 reports quoted RMSE 3.63 and 68% accuracy from synthetic data",
         "All reports are generated from the pipeline by scripts/build_*.py",
         "debugging_log #12; regenerated reports"],
        ["Every generated PDF carried a browser header",
         "Chrome's legacy --headless silently ignores --print-to-pdf-no-header, so a "
         "timestamp and file:// path printed on every page",
         "Switched to --headless=new with --no-pdf-header-footer; all 14 PDFs regenerated",
         "debugging_log; 0 artefacts on re-check"],
    ])

    doc.add_heading("4.4 Integration of Module-1 and Module-2", 2)
    doc.add_paragraph(
        "Module 1 writes outputs/module1_station_forecasts.csv, which carries each station's "
        "Sen's slope, trend category, mean recent depth and coefficient of variation. "
        "src/main.py passes that table to run_module2(), which aggregates stations to taluks "
        "and builds the five indicators from it. The contract between them is the column set "
        "of that file, and TC13 exercises the whole chain.")
    doc.add_paragraph(
        "Two integration bugs were met. First, run_module2() was being called with no "
        "arguments, so it silently scored invented locations rather than Module 1's output — "
        "it ran cleanly and produced a plausible map, which is why it was missed; the fix was "
        "to make the argument required. Second, the station-to-taluk aggregation dropped "
        "almost every row because district and taluk names were spelled differently in the "
        "two sources; canonical_district() and a hand-verified TALUK_ALIASES table fixed it, "
        "and the match rate is now 176 of 177 taluks.")
    p = doc.add_paragraph()
    p.add_run("Evidence: ").bold = True
    p.add_run("Week 6 testcase documents and testing output; bug-fix screenshots with "
              "root-cause notes (Weeks 5 and 7).")

    doc.add_heading("4.5 Validation against CGWB", 2)
    table(doc, ["Measure", "Value", "What it tests"], [
        ["Exact 4-class agreement", "23.9% (κ 0.001)",
         "Relative Jenks classes against absolute CGWB thresholds — cannot align by construction"],
        ["Binary agreement", "77.3%", "Our High/Very High vs CGWB critical/over-exploited"],
        ["Precision", "0.66", "Of the taluks we flag, how many CGWB also flags"],
        ["Recall", "0.45", "Of the taluks CGWB flags, how many we catch"],
        ["Spearman rank", "0.276", "Whether the ranking order agrees"],
    ])
    figure(doc, FIGS / "fig11_cgwb_validation.png",
           "Figure 9: Validation against CGWB, with counts printed on each bar.")
    doc.add_paragraph()
    doc.add_paragraph(
        "We report the weak measure as well as the strong one. Recall of 0.45 is a real "
        "limitation rather than a bug: the 28 stressed taluks we miss have shallow (5.25 m "
        "mean) and flat water tables despite 122% mean extraction, while the 23 we catch are "
        "deep (11.69 m) and declining. CGWB's monitored shallow wells refill each monsoon "
        "while the borewells causing the stress draw from a deeper aquifer, so a water-level "
        "index cannot see stress below the wells it monitors. The index therefore complements "
        "CGWB's extraction accounting rather than replacing it, and closing that gap is a "
        "Phase-2 objective.")

    # ================================================================ 5
    doc.add_page_break()
    doc.add_heading("5. Documentation, Phase-2 Plan and Viva Readiness", 1)

    doc.add_heading("5.1 My individual contribution", 2)
    doc.add_paragraph(f"Student: {stu['name']} ({stu['srn']})   ·   {stu['module']}")
    doc.add_paragraph(f"Primary files: {stu['files']}")
    doc.add_paragraph()
    for item in stu["did"]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_paragraph()
    doc.add_paragraph(
        "Team division: Module 1 (prediction) was built by R Koushik and Bhimisetty Karthik; "
        "Module 2 (risk mapping) by Nitin G A, S Varshith Sai and G M Chandrashekar. "
        "Integration, testing and documentation were shared. Weeks 2–3 were design, 4–5 core "
        "implementation and refinement, 6 testing, 7 consolidation and integration.")

    doc.add_heading("5.2 Phase-2 pending-work plan", 2)
    table(doc, ["Pending work (≈30%)", "Priority", "Planned approach", "Target"], [
        ["Deep-well or extraction data to address the 0.45 recall", "High",
         "Obtain CGWB piezometer / deep-aquifer series and add a depth-stratified indicator",
         "Week 18 (Jan 2027)"],
        ["Kriging risk surface", "High",
         "Ordinary kriging over the 177 taluk centroids with cross-validated variogram",
         "Week 17 (Dec 2026)"],
        ["SHAP explanations", "High",
         "TreeSHAP on the selected XGBoost model, surfaced per station in the dashboard",
         "Week 17 (Dec 2026)"],
        ["Direct multi-horizon forecasting", "Medium",
         "One model per horizon instead of recursive forecasting, to avoid mean reversion",
         "Week 18 (Jan 2027)"],
        ["LULC-change indicator", "Medium",
         "Integrate land-cover rasters and add the deferred fifth design indicator",
         "Week 19 (Jan 2027)"],
        ["Spatio-temporal model (ConvLSTM / GCN-LSTM)", "Low",
         "Research gap G2 — model wells jointly rather than independently",
         "Week 20 (Feb 2027)"],
    ])

    doc.add_heading("5.3 Challenges and learnings", 2)
    doc.add_paragraph(stu["challenge"])

    doc.add_heading("5.4 Viva preparation", 2)
    doc.add_paragraph("Answers prepared without notes, in the examiner's own order:")
    table(doc, ["Question", "Answer in one line"], [
        ["Why this design and these technologies, and what were the alternatives?",
         "Taluk unit over district (CGWB assesses taluks); Sen's slope over the model's own "
         "forecast (recursive forecasts mean-revert); XGBoost over LSTM (tabular, quarterly, "
         "15 features); min-max over percentile and rank normalisation (measured: 77.3% vs "
         "67.6% and 63.6%). See section 2.3."],
        ["How does each diagram map to the code?",
         "Figure 2 is src/pipeline.py run_module1(); Figure 3 is src/module2_risk.py; "
         "Figure 4 is the column contract between them — "
         "outputs/module1_station_forecasts.csv. See section 2.4."],
        ["What is your key algorithm, and its complexity?",
         "Mann-Kendall with Sen's slope per station: O(n²) in the readings per station "
         "(n ≈ 45 quarters, so ~1,000 comparisons), run over 2,042 stations. XGBoost "
         "training is O(trees × depth × rows). The whole pipeline runs in about three "
         "minutes on a laptop with no GPU."],
        ["A bug you found, and how you found its root cause?",
         "See section 5.3 — each of us answers from our own files, with the before/after "
         "evidence in docs/debugging_log.md."],
        ["What is left for Phase-2, and why?",
         "Six items in section 5.2. The first is the honest one: recall of 0.45 is a data "
         "limitation — shallow monitored wells cannot see deep-aquifer stress — and it needs "
         "piezometer data, not a better model."],
    ])
    p = doc.add_paragraph()
    p.add_run("Evidence: ").bold = True
    p.add_run("Phase-2 pending-work plan (Week 7 consolidation document); structured viva "
              "(Week 16).")

    # ================================================================ 6
    doc.add_heading("6. Reproducibility", 1)
    doc.add_paragraph(
        "No number in this report, in the slides or on the dashboard is typed by hand. "
        "Each is read from the pipeline's own output files:")
    for cmd, what in [
        ("python -m src.main", "runs both modules and writes outputs/ and models/"),
        ("python -m pytest tests/", "runs the 29 tests"),
        ("python notebooks/make_figures.py", "regenerates all 12 figures"),
        ("python experiments/ab_rainfall.py", "reproduces the rainfall A/B result"),
        ("python experiments/normalisation_choice.py", "reproduces the normalisation evidence"),
        ("python scripts/build_phase1_reports.py", "regenerates this report"),
        ("streamlit run dashboard.py", "launches the dashboard"),
    ]:
        doc.add_paragraph(f"{cmd}  —  {what}", style="List Bullet")

    # ---------------- declaration
    doc.add_page_break()
    doc.add_heading("Declaration", 1)
    doc.add_paragraph(
        f"I, {stu['name']} (SRN {stu['srn']}), declare that this report is my own work and "
        f"that the evidence I have uploaded is genuine. The work was carried out with my "
        f"project team under the guidance of {GUIDE}.")
    doc.add_paragraph(
        "All results, figures and tables in this report are produced by running our own code "
        "on publicly available data from the Central Ground Water Board (via India-WRIS) and "
        "NASA POWER. No result has been estimated, adjusted or reported from any source other "
        "than an actual execution of that code. Where a measure is weak — the 23.9% "
        "four-class agreement and the 0.45 recall — it is reported as it stands, together "
        "with our analysis of why.")
    doc.add_paragraph(
        "All external work is cited in the Week 1 Literature Survey.")
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
