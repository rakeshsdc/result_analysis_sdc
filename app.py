"""
FYUGP / University Mark-list Result Analysis App
--------------------------------------------------
Upload a University "Mark cum Grade Statement" PDF (one page per student,
in the standard University of Kerala FYUGP layout) and generate a full
Result Analysis Report — viewable on screen and downloadable as a
formatted Word (.docx) document with a college-logo letterhead.

Run locally:
    streamlit run app.py

Deploy:
    Push this folder to a GitHub repo (app.py + requirements.txt) and
    deploy on https://share.streamlit.io (Streamlit Community Cloud),
    pointing it at app.py.
"""

import io
import re
from collections import Counter

import pandas as pd
import pdfplumber
import matplotlib.pyplot as plt
import streamlit as st

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --------------------------------------------------------------------------
# The official grade hierarchy used by the university (best -> worst).
# "F" here represents an overall Fail result (the mark sheet prints "--"
# for Grade Awarded when a student fails, so we relabel it "F" for the
# purposes of grouping/ordering in this report).
# --------------------------------------------------------------------------
GRADE_ORDER = ["O", "A+", "A", "B+", "B", "C", "P", "F"]

# --------------------------------------------------------------------------
# Fixed institution identity (no longer user-editable).
# --------------------------------------------------------------------------
import os

COLLEGE_NAME = "Sanatana Dharma College, Alappuzha"
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo.jpg")

COPYRIGHT_NOTICE = (
    "Developed by Dr. Rakesh Chandran S. B., Associate Professor, "
    "Department of Physics, Sanatana Dharma College, Alappuzha. "
    "An IQAC SD College initiative."
)


def _load_logo_bytes():
    try:
        with open(LOGO_PATH, "rb") as f:
            return f.read()
    except FileNotFoundError:
        return None


def _grade_sort_key(g):
    return GRADE_ORDER.index(g) if g in GRADE_ORDER else len(GRADE_ORDER)


def effective_grade(student):
    """Grade Awarded if present, else 'F' if the overall result is Fail."""
    if student["grade_awarded"]:
        return student["grade_awarded"]
    if student["status"].lower() == "fail":
        return "F"
    return None


# --------------------------------------------------------------------------
# ---------------------------  PDF PARSING  -------------------------------
# --------------------------------------------------------------------------

def _to_float(val):
    """Safely convert a mark-sheet cell to float, else return None."""
    if val is None:
        return None
    val = str(val).strip()
    if val in ("", "*", "-", "--", "Ab"):
        return None
    try:
        return float(val)
    except ValueError:
        return None


def parse_marklist_pdf(file_bytes):
    """
    Parse the uploaded PDF (bytes) and return:
        students      : list of dicts, one per student
        course_names  : dict {course_code: course_title}
    """
    students = []
    course_names = {}

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            m_id = re.search(r"APAAR ID\s*:\s*([0-9A-Za-z]+)", text)
            m_name = re.search(r"Name of Student\s*:\s*(.+)", text)
            m_prog = re.search(r"Name of Programme\s*:\s*(.+)", text)

            if not m_id or not m_name:
                continue

            apaar = m_id.group(1).strip()
            name = m_name.group(1).strip()
            programme = m_prog.group(1).strip() if m_prog else ""

            tables = page.extract_tables()
            if not tables:
                continue
            main_table = tables[0]

            courses = []
            total_row = None
            result_row = None

            for row in main_table:
                first = (row[0] or "").strip() if row[0] else ""

                if first == "Total":
                    total_row = row
                    continue
                if first.startswith("Result:"):
                    result_row = row
                    continue
                if row[1] is None or first in ("", "Course Code"):
                    continue

                code = first
                title = (row[1] or "").replace("\n", " ").strip()
                if not code or not title:
                    continue

                credits = row[2]
                course_total = _to_float(row[13])
                gp = _to_float(row[14])
                cp = _to_float(row[15])
                grade = (row[16] or "").strip()

                course_names.setdefault(code, title)
                courses.append(
                    {
                        "code": code,
                        "title": title,
                        "credits": credits,
                        "course_total": course_total,
                        "gp": gp,
                        "cp": cp,
                        "grade": grade,
                    }
                )

            total_credits = ""
            total_cp = None
            if total_row:
                total_credits = (total_row[2] or "").strip()
                for idx in (15, 14, 13):
                    val = total_row[idx]
                    if val:
                        total_cp = _to_float(val)
                        break

            status, sgpa, grade_awarded = "Unknown", None, ""
            if result_row:
                joined = " ".join([c for c in result_row if c])
                m_status = re.search(r"Result\s*:\s*(Pass|Fail)", joined, re.I)
                m_sgpa = re.search(r"SGPA\s*:\s*([\d.]+|--)", joined)
                m_grade = re.search(r"Grade awarded\s*:\s*([A-Za-z+\-]+|--)", joined)
                if m_status:
                    status = m_status.group(1).capitalize()
                if m_sgpa and m_sgpa.group(1) != "--":
                    sgpa = float(m_sgpa.group(1))
                if m_grade and m_grade.group(1) != "--":
                    grade_awarded = m_grade.group(1)

            total_marks = sum(c["course_total"] for c in courses if c["course_total"])
            total_marks = round(total_marks, 2)

            students.append(
                {
                    "name": name,
                    "apaar": apaar,
                    "programme": programme,
                    "courses": courses,
                    "total_credits": total_credits,
                    "total_cp": total_cp,
                    "sgpa": sgpa,
                    "grade_awarded": grade_awarded,
                    "status": status,
                    "total_marks": total_marks,
                }
            )

    return students, course_names


# --------------------------------------------------------------------------
# ---------------------------  ANALYSIS  -----------------------------------
# --------------------------------------------------------------------------

def build_student_table(students):
    rows = []
    for i, s in enumerate(students, start=1):
        rows.append(
            {
                "Sl.No": i,
                "Name": s["name"],
                "APAAR ID": s["apaar"],
                "Total Credits": s["total_credits"] if s["total_credits"] else "--",
                "SGPA": s["sgpa"] if s["sgpa"] is not None else "--",
                "Grade Awarded": s["grade_awarded"] if s["grade_awarded"] else "--",
                "Status": s["status"],
            }
        )
    return pd.DataFrame(rows)


def build_result_summary(students):
    appeared = len(students)
    passed = sum(1 for s in students if s["status"].lower() == "pass")
    failed = sum(1 for s in students if s["status"].lower() == "fail")
    other = appeared - passed - failed
    pass_pct = round((passed / appeared) * 100, 2) if appeared else 0.0

    summary_df = pd.DataFrame(
        [
            {"Category": "Appeared", "Count": appeared},
            {"Category": "Passed", "Count": passed},
            {"Category": "Failed", "Count": failed},
        ]
        + ([{"Category": "Result Not Available", "Count": other}] if other else [])
    )
    return summary_df, pass_pct


def build_top_rankers(students, top_n=3):
    ranked = [s for s in students if s["sgpa"] is not None]
    if not ranked:
        return pd.DataFrame()
    ranked.sort(key=lambda s: s["sgpa"], reverse=True)

    distinct_sgpas = sorted({s["sgpa"] for s in ranked}, reverse=True)
    cutoff_values = distinct_sgpas[:top_n]

    rows = []
    for pos, val in enumerate(cutoff_values, start=1):
        holders = [s for s in ranked if s["sgpa"] == val]
        for s in holders:
            rows.append(
                {
                    "Rank": pos,
                    "Name": s["name"],
                    "APAAR ID": s["apaar"],
                    "SGPA": s["sgpa"],
                    "Grade Awarded": s["grade_awarded"],
                }
            )
    return pd.DataFrame(rows)


def build_subject_failure_table(students, course_names):
    fail_counts = {code: 0 for code in course_names}
    appeared_counts = {code: 0 for code in course_names}

    for s in students:
        for c in s["courses"]:
            code = c["code"]
            appeared_counts[code] = appeared_counts.get(code, 0) + 1
            if c["grade"].upper() == "F":
                fail_counts[code] = fail_counts.get(code, 0) + 1

    rows = []
    for code in course_names:
        rows.append(
            {
                "Course Code": code,
                "Course Title": course_names[code],
                "Appeared": appeared_counts.get(code, 0),
                "Failed": fail_counts.get(code, 0),
            }
        )
    df = pd.DataFrame(rows).sort_values("Failed", ascending=False).reset_index(drop=True)
    return df


def build_grade_distribution(students):
    """Grade Awarded vs. No. of Students, ordered O, A+, A, B+, B, C, P, F."""
    counts = Counter()
    for s in students:
        g = effective_grade(s)
        if g:
            counts[g] += 1
    ordered_grades = sorted(counts.keys(), key=_grade_sort_key)
    rows = [{"Grade Awarded": g, "No. of Students": counts[g]} for g in ordered_grades]
    return pd.DataFrame(rows)


def build_grade_wise_student_list(students):
    """Sl.No, Name, Grade — rows grouped together by grade, in hierarchy order."""
    groups = {}
    for s in students:
        g = effective_grade(s) or "NA"
        groups.setdefault(g, []).append(s)

    ordered_grades = sorted(groups.keys(), key=_grade_sort_key)
    rows = []
    sl_no = 1
    for g in ordered_grades:
        for s in sorted(groups[g], key=lambda x: x["name"]):
            rows.append({"Sl.No": sl_no, "Name": s["name"], "Grade": g})
            sl_no += 1
    return pd.DataFrame(rows), ordered_grades, groups


def build_summary_paragraphs(
    exam_details,
    exam_month_year,
    programme_name,
    summary_df,
    pass_pct,
    top_df,
    fail_df,
    avg_sgpa,
    highest_sgpa,
    lowest_sgpa,
):
    """Return a list of paragraph strings summarising the whole report."""
    def _count(cat):
        rows = summary_df.loc[summary_df["Category"] == cat, "Count"]
        return int(rows.iloc[0]) if not rows.empty else 0

    appeared = _count("Appeared")
    passed = _count("Passed")
    failed = _count("Failed")

    paras = []

    exam_line = exam_details
    if exam_month_year:
        exam_line = f"{exam_details} ({exam_month_year})" if exam_details else exam_month_year

    paras.append(
        f"A total of {appeared} student(s) appeared for the {exam_line} "
        f"held for the {programme_name} programme. Of these, {passed} "
        f"student(s) passed and {failed} student(s) failed, resulting in "
        f"an overall pass percentage of {pass_pct}%."
    )

    if avg_sgpa is not None:
        paras.append(
            f"Among the students who passed, the average SGPA secured was "
            f"{avg_sgpa}, with the highest SGPA recorded at {highest_sgpa} "
            f"and the lowest passing SGPA at {lowest_sgpa}."
        )

    if top_df is not None and not top_df.empty:
        top_rank1 = top_df[top_df["Rank"] == 1]
        names = ", ".join(top_rank1["Name"].tolist())
        sgpa_val = top_rank1["SGPA"].iloc[0]
        verb = "secured" if len(top_rank1) == 1 else "jointly secured"
        paras.append(f"{names} {verb} the top rank with an SGPA of {sgpa_val}.")

    if fail_df is not None and not fail_df.empty and fail_df["Failed"].sum() > 0:
        top_fail = fail_df.iloc[0]
        paras.append(
            f"The highest number of failures was recorded in "
            f"{top_fail['Course Code']} ({top_fail['Course Title']}), where "
            f"{top_fail['Failed']} student(s) failed — indicating this "
            f"course may need additional academic support."
        )
    else:
        paras.append("No student failed in any individual course.")

    return paras


# --------------------------------------------------------------------------
# ---------------------------  DOCX BUILDING  ------------------------------
# --------------------------------------------------------------------------

def _set_cell_shading(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def _add_bottom_border(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "auto")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _add_heading(doc, text, size=13):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    return p


def _add_table_from_df(doc, df, header_color="4472C4", header_font_color="FFFFFF"):
    if df is None or df.empty:
        doc.add_paragraph("(No data)")
        return
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    for i, col in enumerate(df.columns):
        hdr_cells[i].text = str(col)
        for p in hdr_cells[i].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.color.rgb = RGBColor.from_string(header_font_color)
        _set_cell_shading(hdr_cells[i], header_color)

    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, col in enumerate(df.columns):
            cells[i].text = "" if pd.isna(row[col]) else str(row[col])
            for p in cells[i].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return table


def _fig_to_docx(doc, fig, width_inches=5.5):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight")
    buf.seek(0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(buf, width=Inches(width_inches))


def _add_signature_block(doc):
    """Three signature slots — Faculty Advisor, HOD, Principal — in one row."""
    # Leave some blank vertical space for the actual pen-and-ink signature.
    doc.add_paragraph()
    doc.add_paragraph()

    labels = ["Faculty Advisor", "Head of the Department", "Principal"]

    table = doc.add_table(rows=2, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    col_width = Inches(2.1)
    for row in table.rows:
        for cell in row.cells:
            cell.width = col_width
    for col in table.columns:
        col.width = col_width

    for i, label in enumerate(labels):
        line_p = table.cell(0, i).paragraphs[0]
        line_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        line_p.add_run("_" * 24)

        label_p = table.cell(1, i).paragraphs[0]
        label_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        label_run = label_p.add_run(label)
        label_run.bold = True
        label_run.font.size = Pt(11)


def generate_docx_report(
    college_name,
    exam_details,
    exam_month_year,
    programme_name,
    logo_bytes,
    student_df,
    summary_df,
    pass_pct,
    top_df,
    fail_df,
    legend_df,
    grade_dist_df,
    grade_wise_df,
    avg_sgpa=None,
    highest_sgpa=None,
    lowest_sgpa=None,
    fig_summary=None,
    fig_fail=None,
    fig_grade=None,
):
    doc = Document()

    # ---------------- Letterhead / Header ----------------
    if logo_bytes:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(io.BytesIO(logo_bytes), width=Inches(0.9))

    if college_name:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(college_name)
        run.bold = True
        run.font.size = Pt(16)

    if exam_details or exam_month_year:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if exam_details and exam_month_year:
            line = f"{exam_details}, {exam_month_year}"
        else:
            line = exam_details or exam_month_year
        run = p.add_run(line)
        run.font.size = Pt(12)

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Result Analysis Report")
    run.bold = True
    run.underline = True
    run.font.size = Pt(14)

    if programme_name:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = p2.add_run(f"Programme: {programme_name}")
        run2.font.size = Pt(12)

    hr_p = doc.add_paragraph()
    _add_bottom_border(hr_p)

    # ---------------- 1. Consolidated Result Table ----------------
    _add_heading(doc, "1. Consolidated Result Table")
    _add_table_from_df(doc, student_df)

    # ---------------- 2. Result Summary ----------------
    _add_heading(doc, "2. Result Summary")
    _add_table_from_df(doc, summary_df)
    p = doc.add_paragraph()
    run = p.add_run(f"Pass Percentage: {pass_pct}%")
    run.bold = True
    if fig_summary is not None:
        _fig_to_docx(doc, fig_summary, width_inches=4.5)

    # ---------------- 3. Top Rank Holders ----------------
    _add_heading(doc, "3. Top Rank Holders")
    _add_table_from_df(doc, top_df)

    # ---------------- 4. Subject-wise Failure Analysis ----------------
    _add_heading(doc, "4. Subject-wise Failure Analysis")
    _add_table_from_df(doc, fail_df)
    if not fail_df.empty and fail_df["Failed"].sum() > 0:
        top_row = fail_df.iloc[0]
        p = doc.add_paragraph()
        run = p.add_run(
            f"Highest number of failures: {top_row['Course Code']} — "
            f"{top_row['Course Title']} ({top_row['Failed']} student(s))"
        )
        run.italic = True
    if fig_fail is not None:
        _fig_to_docx(doc, fig_fail, width_inches=5.5)

    # ---------------- 5. Course Code / Full Name Reference ----------------
    _add_heading(doc, "5. Course Code / Full Name Reference")
    _add_table_from_df(doc, legend_df)

    # ---------------- 6. Grade Distribution ----------------
    _add_heading(doc, "6. Grade Distribution")
    sgpa_stats_df = pd.DataFrame(
        [
            {"Metric": "Average SGPA (Passed Students)", "Value": avg_sgpa if avg_sgpa is not None else "--"},
            {"Metric": "Highest SGPA", "Value": highest_sgpa if highest_sgpa is not None else "--"},
            {"Metric": "Lowest SGPA (among passed)", "Value": lowest_sgpa if lowest_sgpa is not None else "--"},
        ]
    )
    _add_table_from_df(doc, sgpa_stats_df)
    doc.add_paragraph()
    _add_table_from_df(doc, grade_dist_df)
    if fig_grade is not None:
        _fig_to_docx(doc, fig_grade, width_inches=4.5)

    # ---------------- 7. Grade-wise List of Students ----------------
    _add_heading(doc, "7. Grade-wise List of Students")
    _add_table_from_df(doc, grade_wise_df)

    # ---------------- 8. Summary ----------------
    _add_heading(doc, "8. Summary")
    summary_paragraphs = build_summary_paragraphs(
        exam_details=exam_details,
        exam_month_year=exam_month_year,
        programme_name=programme_name,
        summary_df=summary_df,
        pass_pct=pass_pct,
        top_df=top_df,
        fail_df=fail_df,
        avg_sgpa=avg_sgpa,
        highest_sgpa=highest_sgpa,
        lowest_sgpa=lowest_sgpa,
    )
    for para_text in summary_paragraphs:
        p = doc.add_paragraph(para_text)
        p.paragraph_format.space_after = Pt(8)

    # ---------------- Signatures ----------------
    _add_signature_block(doc)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------------
# ---------------------------  STREAMLIT UI  -------------------------------
# --------------------------------------------------------------------------

st.set_page_config(page_title="Result Analysis Report", layout="wide")

st.title("📊 Result Analysis Report Generator")
st.caption(
    "Upload a University mark-cum-grade-statement PDF (one page per student) "
    "and generate a consolidated result analysis report."
)

college_name = COLLEGE_NAME
logo_bytes = _load_logo_bytes()

with st.sidebar:
    st.header("Report Details")
    st.markdown(f"**Institution:** {COLLEGE_NAME}")
    exam_details = st.text_input(
        "Examination Details", "Fourth Semester FYUGP Examination"
    )
    programme_name = st.text_input("Programme / Major", "Physics")
    exam_month_year = st.text_input("Month & Year", "May 2026")

    st.markdown("---")
    uploaded_pdf = st.file_uploader("Upload Mark-list PDF", type=["pdf"])
    generate = st.button("Generate Report", type="primary", use_container_width=True)

if generate and uploaded_pdf is not None:
    with st.spinner("Reading and parsing the PDF..."):
        students, course_names = parse_marklist_pdf(uploaded_pdf.read())

    if not students:
        st.error(
            "Could not find any student mark statements in this PDF. "
            "Please check the file and try again."
        )
        st.stop()

    st.session_state["students"] = students
    st.session_state["course_names"] = course_names

if "students" in st.session_state:
    students = st.session_state["students"]
    course_names = st.session_state["course_names"]

    # ---------------- Report Header ----------------
    st.markdown("---")
    header_html = "<div style='text-align:center'>"
    if logo_bytes:
        import base64
        b64 = base64.b64encode(logo_bytes).decode()
        header_html += f"<img src='data:image/png;base64,{b64}' style='height:70px'><br>"
    if college_name:
        header_html += f"<h2 style='margin-bottom:0'>{college_name}</h2>"
    header_html += f"<h4 style='margin-top:4px'>{exam_details} &nbsp;&mdash;&nbsp; {exam_month_year}</h4>"
    header_html += f"<h3 style='margin-top:4px;text-decoration:underline'>Result Analysis Report</h3>"
    header_html += f"<p style='font-size:16px'><b>Programme:</b> {programme_name}</p>"
    header_html += "</div>"
    st.markdown(header_html, unsafe_allow_html=True)
    st.markdown("---")

    # ---------------- 1. Consolidated Student Table ----------------
    st.subheader("1. Consolidated Result Table")
    student_df = build_student_table(students)
    st.dataframe(student_df, use_container_width=True, hide_index=True)

    # ---------------- 2. Result Summary ----------------
    st.subheader("2. Result Summary")
    summary_df, pass_pct = build_result_summary(students)
    col1, col2 = st.columns([1, 1.4])
    fig_summary = None
    with col1:
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
        st.metric("Pass Percentage", f"{pass_pct}%")
    with col2:
        fig_summary, ax = plt.subplots(figsize=(4, 3))
        colors = {"Appeared": "#4C72B0", "Passed": "#55A868", "Failed": "#C44E52",
                  "Result Not Available": "#8172B2"}
        bar_colors = [colors.get(c, "#999999") for c in summary_df["Category"]]
        ax.bar(summary_df["Category"], summary_df["Count"], color=bar_colors)
        ax.set_ylabel("No. of Students")
        ax.set_title("Appeared / Passed / Failed")
        for i, v in enumerate(summary_df["Count"]):
            ax.text(i, v + 0.1, str(v), ha="center", fontweight="bold")
        plt.xticks(rotation=15)
        st.pyplot(fig_summary, use_container_width=True)

    # ---------------- 3. Top Rank Holders ----------------
    st.subheader("3. Top Rank Holders")
    top_df = build_top_rankers(students, top_n=3)
    if top_df.empty:
        st.info("No student has a valid SGPA to rank (all results unavailable).")
    else:
        st.dataframe(top_df, use_container_width=True, hide_index=True)

    # ---------------- 4. Subject-wise Failure Analysis ----------------
    st.subheader("4. Subject-wise Failure Analysis")
    fail_df = build_subject_failure_table(students, course_names)
    st.dataframe(fail_df, use_container_width=True, hide_index=True)

    fig_fail = None
    if fail_df["Failed"].sum() > 0:
        top_fail_row = fail_df.iloc[0]
        st.warning(
            f"📌 Highest number of failures is in **{top_fail_row['Course Code']} "
            f"— {top_fail_row['Course Title']}** with **{top_fail_row['Failed']}** "
            "student(s) failing."
        )
        fig_fail, ax2 = plt.subplots(figsize=(8, 4))
        plot_df = fail_df[fail_df["Failed"] > 0]
        ax2.bar(plot_df["Course Code"], plot_df["Failed"], color="#C44E52")
        ax2.set_ylabel("No. of Students Failed")
        ax2.set_title("Failures by Course")
        for i, v in enumerate(plot_df["Failed"]):
            ax2.text(i, v + 0.05, str(v), ha="center", fontweight="bold")
        plt.xticks(rotation=30, ha="right")
        st.pyplot(fig_fail, use_container_width=True)
    else:
        st.success("No student failed in any subject. 🎉")

    # ---------------- 5. Course Code -> Full Name Legend ----------------
    st.subheader("5. Course Code / Full Name Reference")
    legend_df = pd.DataFrame(
        [{"Course Code": k, "Course Full Name": v} for k, v in course_names.items()]
    ).sort_values("Course Code").reset_index(drop=True)
    st.dataframe(legend_df, use_container_width=True, hide_index=True)

    # ---------------- 6. Grade Distribution ----------------
    st.subheader("6. Grade Distribution")
    grade_dist_df = build_grade_distribution(students)
    valid_sgpas = [s["sgpa"] for s in students if s["sgpa"] is not None]
    avg_sgpa = round(sum(valid_sgpas) / len(valid_sgpas), 2) if valid_sgpas else None
    highest_sgpa = max(valid_sgpas) if valid_sgpas else None
    lowest_sgpa = min(valid_sgpas) if valid_sgpas else None

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Average SGPA (Passed Students)", avg_sgpa if avg_sgpa is not None else "--")
    with c2:
        st.metric("Highest SGPA", highest_sgpa if highest_sgpa is not None else "--")
    with c3:
        st.metric("Lowest SGPA (among passed)", lowest_sgpa if lowest_sgpa is not None else "--")

    fig_grade = None
    if not grade_dist_df.empty:
        colA, colB = st.columns([1, 1.4])
        with colA:
            st.dataframe(grade_dist_df, use_container_width=True, hide_index=True)
        with colB:
            fig_grade, ax3 = plt.subplots(figsize=(4, 3))
            ax3.bar(grade_dist_df["Grade Awarded"], grade_dist_df["No. of Students"], color="#4C72B0")
            ax3.set_title("Grade Distribution")
            ax3.set_ylabel("No. of Students")
            st.pyplot(fig_grade, use_container_width=True)

    # ---------------- 7. Grade-wise List of Students ----------------
    st.subheader("7. Grade-wise List of Students")
    grade_wise_df, ordered_grades, groups = build_grade_wise_student_list(students)
    st.dataframe(grade_wise_df, use_container_width=True, hide_index=True)

    # ---------------- 8. Summary ----------------
    st.subheader("8. Summary")
    summary_paragraphs = build_summary_paragraphs(
        exam_details=exam_details,
        exam_month_year=exam_month_year,
        programme_name=programme_name,
        summary_df=summary_df,
        pass_pct=pass_pct,
        top_df=top_df,
        fail_df=fail_df,
        avg_sgpa=avg_sgpa,
        highest_sgpa=highest_sgpa,
        lowest_sgpa=lowest_sgpa,
    )
    for para_text in summary_paragraphs:
        st.write(para_text)

    st.markdown("#####")
    sig1, sig2, sig3 = st.columns(3)
    for col, label in zip((sig1, sig2, sig3), ("Faculty Advisor", "Head of the Department", "Principal")):
        with col:
            st.markdown("&nbsp;")
            st.markdown("____________________")
            st.markdown(f"**{label}**")

    st.markdown("---")
    st.caption(
        "Note: 'Total Credits' and 'SGPA' are taken as printed on the "
        "official mark statement (shown as '--' where the university's "
        "statement itself shows '--', e.g. when the result is Fail). "
        "In the Grade Distribution and Grade-wise List, students whose "
        "overall result is Fail are grouped under grade 'F'."
    )

    # ---------------- DOCX download ----------------
    st.markdown("---")
    docx_bytes = generate_docx_report(
        college_name=college_name,
        exam_details=exam_details,
        exam_month_year=exam_month_year,
        programme_name=programme_name,
        logo_bytes=logo_bytes,
        student_df=student_df,
        summary_df=summary_df,
        pass_pct=pass_pct,
        top_df=top_df,
        fail_df=fail_df,
        legend_df=legend_df,
        grade_dist_df=grade_dist_df,
        grade_wise_df=grade_wise_df,
        avg_sgpa=avg_sgpa,
        highest_sgpa=highest_sgpa,
        lowest_sgpa=lowest_sgpa,
        fig_summary=fig_summary,
        fig_fail=fig_fail,
        fig_grade=fig_grade,
    )
    st.download_button(
        "📄 Download Full Report (.docx)",
        docx_bytes,
        file_name="Result_Analysis_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )
    st.download_button(
        "⬇️ Download Student Table Only (CSV)",
        student_df.to_csv(index=False).encode("utf-8"),
        file_name="student_result_table.csv",
        mime="text/csv",
    )

    st.markdown("---")
    st.caption(f"© {COPYRIGHT_NOTICE}")

elif not uploaded_pdf:
    st.info("⬅️ Fill in the examination details and upload the mark-list PDF from the sidebar, then click **Generate Report**.")
