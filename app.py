"""
FYUGP / University Mark-list Result Analysis App
--------------------------------------------------
Upload a University "Mark cum Grade Statement" PDF (one page per student,
in the standard University of Kerala FYUGP layout) and generate a full
Result Analysis Report:

  * Consolidated table of all students (Name, APAAR ID, Total Marks,
    Total Credits, SGPA, Grade Awarded, Status)
  * Appeared / Passed / Failed summary + pass percentage + bar chart
  * Top-3 rank holders (ties included)
  * Subject-wise (course-wise) failure counts + bar chart + course-name
    legend table
  * A few extra analyses: grade distribution, average SGPA, highest /
    lowest SGPA

Run locally:
    streamlit run app.py

Deploy:
    Push this folder to a GitHub repo (app.py + requirements.txt) and
    deploy on https://share.streamlit.io (Streamlit Community Cloud),
    pointing it at app.py.
"""

import io
import re

import pandas as pd
import pdfplumber
import matplotlib.pyplot as plt
import streamlit as st

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

            # Skip pages that are not a student mark statement
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
                # A real course row always has a Course Title (row[1]);
                # section separators (e.g. "DSC 10-Major Paper 4", "SEC-1")
                # and the header rows have row[1] == None.
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

            # ---- Total row (credits earned, total CP) ----
            total_credits = ""
            total_cp = None
            if total_row:
                total_credits = (total_row[2] or "").strip()
                for idx in (15, 14, 13):
                    val = total_row[idx]
                    if val:
                        total_cp = _to_float(val)
                        break

            # ---- Result row (Result / SGPA / Grade awarded) ----
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

            # ---- Derived: total marks actually obtained (sum of course totals) ----
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
                "Total Marks": s["total_marks"],
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
                    "Total Marks": s["total_marks"],
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
    grades = [s["grade_awarded"] for s in students if s["grade_awarded"]]
    if not grades:
        return pd.DataFrame()
    series = pd.Series(grades).value_counts().reset_index()
    series.columns = ["Grade Awarded", "No. of Students"]
    return series


# --------------------------------------------------------------------------
# ---------------------------  STREAMLIT UI  -------------------------------
# --------------------------------------------------------------------------

st.set_page_config(page_title="Result Analysis Report", layout="wide")

st.title("📊 Result Analysis Report Generator")
st.caption(
    "Upload a University mark-cum-grade-statement PDF (one page per student) "
    "and generate a consolidated result analysis report."
)

with st.sidebar:
    st.header("Report Details")
    college_name = st.text_input("College / Institution Name", "")
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
    if college_name:
        header_html += f"<h2 style='margin-bottom:0'>{college_name}</h2>"
    header_html += f"<h4 style='margin-top:4px'>{exam_details} &nbsp;&mdash;&nbsp; {exam_month_year}</h4>"
    header_html += f"<h3 style='margin-top:4px'>Result Analysis Report</h3>"
    header_html += f"<p style='font-size:16px'><b>Programme:</b> {programme_name}</p>"
    header_html += "</div>"
    st.markdown(header_html, unsafe_allow_html=True)
    st.markdown("---")

    # ---------------- 1. Consolidated Student Table ----------------
    st.subheader("1. Consolidated Result Table")
    student_df = build_student_table(students)
    st.dataframe(student_df, use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Download Student Table (CSV)",
        student_df.to_csv(index=False).encode("utf-8"),
        file_name="student_result_table.csv",
        mime="text/csv",
    )

    # ---------------- 2. Result Summary ----------------
    st.subheader("2. Result Summary")
    summary_df, pass_pct = build_result_summary(students)
    col1, col2 = st.columns([1, 1.4])
    with col1:
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
        st.metric("Pass Percentage", f"{pass_pct}%")
    with col2:
        fig, ax = plt.subplots(figsize=(4, 3))
        colors = {"Appeared": "#4C72B0", "Passed": "#55A868", "Failed": "#C44E52",
                  "Result Not Available": "#8172B2"}
        bar_colors = [colors.get(c, "#999999") for c in summary_df["Category"]]
        ax.bar(summary_df["Category"], summary_df["Count"], color=bar_colors)
        ax.set_ylabel("No. of Students")
        ax.set_title("Appeared / Passed / Failed")
        for i, v in enumerate(summary_df["Count"]):
            ax.text(i, v + 0.1, str(v), ha="center", fontweight="bold")
        plt.xticks(rotation=15)
        st.pyplot(fig, use_container_width=True)

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

    if fail_df["Failed"].sum() > 0:
        top_fail_row = fail_df.iloc[0]
        st.warning(
            f"📌 Highest number of failures is in **{top_fail_row['Course Code']} "
            f"— {top_fail_row['Course Title']}** with **{top_fail_row['Failed']}** "
            "student(s) failing."
        )
        fig2, ax2 = plt.subplots(figsize=(8, 4))
        plot_df = fail_df[fail_df["Failed"] > 0]
        ax2.bar(plot_df["Course Code"], plot_df["Failed"], color="#C44E52")
        ax2.set_ylabel("No. of Students Failed")
        ax2.set_title("Failures by Course")
        for i, v in enumerate(plot_df["Failed"]):
            ax2.text(i, v + 0.05, str(v), ha="center", fontweight="bold")
        plt.xticks(rotation=30, ha="right")
        st.pyplot(fig2, use_container_width=True)
    else:
        st.success("No student failed in any subject. 🎉")

    # ---------------- 5. Course Code -> Full Name Legend ----------------
    st.subheader("5. Course Code / Full Name Reference")
    legend_df = pd.DataFrame(
        [{"Course Code": k, "Course Full Name": v} for k, v in course_names.items()]
    ).sort_values("Course Code").reset_index(drop=True)
    st.dataframe(legend_df, use_container_width=True, hide_index=True)

    # ---------------- 6. Additional Analysis ----------------
    st.subheader("6. Additional Analysis")
    valid_sgpas = [s["sgpa"] for s in students if s["sgpa"] is not None]
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric(
            "Average SGPA (Passed Students)",
            round(sum(valid_sgpas) / len(valid_sgpas), 2) if valid_sgpas else "--",
        )
    with c2:
        st.metric("Highest SGPA", max(valid_sgpas) if valid_sgpas else "--")
    with c3:
        st.metric("Lowest SGPA (among passed)", min(valid_sgpas) if valid_sgpas else "--")

    grade_dist_df = build_grade_distribution(students)
    if not grade_dist_df.empty:
        colA, colB = st.columns([1, 1.4])
        with colA:
            st.dataframe(grade_dist_df, use_container_width=True, hide_index=True)
        with colB:
            fig3, ax3 = plt.subplots(figsize=(4, 3))
            ax3.bar(grade_dist_df["Grade Awarded"], grade_dist_df["No. of Students"], color="#4C72B0")
            ax3.set_title("Grade Distribution")
            ax3.set_ylabel("No. of Students")
            st.pyplot(fig3, use_container_width=True)

    st.markdown("---")
    st.caption(
        "Note: 'Total Marks' is computed as the sum of the Course-Total "
        "(CCA+ESE) marks across all courses shown on each student's mark "
        "statement. 'Total Credits' and 'SGPA' are taken as printed on the "
        "official mark statement (shown as '--' where the university's "
        "statement itself shows '--', e.g. when the result is Fail)."
    )

elif not uploaded_pdf:
    st.info("⬅️ Fill in the report details and upload the mark-list PDF from the sidebar, then click **Generate Report**.")
