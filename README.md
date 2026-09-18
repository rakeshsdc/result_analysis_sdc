# Result Analysis Report Generator

A Streamlit app that takes a University "Mark cum Grade Statement" PDF
(one page per student — the standard University of Kerala FYUGP layout)
and automatically generates a full **Result Analysis Report** for
**Sanatana Dharma College, Alappuzha**, viewable on screen and
downloadable as a formatted **Word (.docx)** document with the
college's logo letterhead.

Report contents:

1. Consolidated table: Sl.No, Name, APAAR ID, Total Credits, SGPA,
   Grade Awarded, Status
2. Appeared / Passed / Failed summary + Pass % + bar chart
3. Top-3 rank holders by SGPA (all tied students included)
4. Subject-wise failure counts (course-code wise) + bar chart, and the
   subject with the highest number of failures is called out
5. Course code → full course name reference table
6. Average / Highest / Lowest SGPA, and Grade Distribution — ordered
   **O, A+, A, B+, B, C, P, F** (the university's own grade hierarchy,
   best to worst; students whose overall result is Fail are counted
   under "F")
7. Grade-wise list of students — Sl.No, Name, Grade — with students
   grouped together by grade, in the same O→F order
8. Summary — an auto-generated narrative covering appeared/passed/
   failed counts, pass %, average/highest/lowest SGPA, the topper(s),
   and the subject with the most failures
9. Signature blocks for **Faculty Advisor**, **Head of the
   Department**, and **Principal**, laid out side by side with a
   signature line above each

The **institution name and logo are fixed** in the app (Sanatana Dharma
College, Alappuzha) — there is nothing to enter for these. The
editable fields in the sidebar are only: **Examination details**
(e.g. "Fourth Semester FYUGP Examination"), **Programme/Major**
(e.g. "Physics"), and **Month & Year** (e.g. "May 2026").

---

## Folder structure

```
fyugp_result_app/
├── app.py
├── requirements.txt
└── assets/
    └── logo.jpg      <- Sanatana Dharma College logo (used in the report header)
```

**Keep `assets/logo.jpg` alongside `app.py`** — the app loads the logo
from that relative path at runtime. If you ever need to change the
logo, just replace this file (keep the same name `logo.jpg`, or update
the `LOGO_PATH` constant at the top of `app.py`).

## 1. Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).

## 2. Deploy for free on Streamlit Community Cloud (via GitHub)

1. Create a new GitHub repository and push the **whole folder**
   (`app.py`, `requirements.txt`, and the `assets/` folder with
   `logo.jpg` inside it):

   ```bash
   git init
   git add app.py requirements.txt assets README.md
   git commit -m "Result analysis app"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```

   > Make sure `assets/logo.jpg` is actually committed — double-check
   > with `git status` before pushing. Without it, the report header
   > will simply skip the logo (everything else still works).

2. Go to **https://share.streamlit.io** and sign in with your GitHub
   account.

3. Click **"New app"**, choose your repository and branch (`main`), and
   set the **Main file path** to `app.py`.

4. Click **Deploy**. Streamlit Cloud installs everything from
   `requirements.txt` automatically (including `python-docx`) and gives
   you a public URL such as `https://<your-app-name>.streamlit.app`.

5. Any time you push new commits to the repo, the deployed app
   auto-updates.

## 3. How to use the app

1. Open the app.
2. In the left sidebar, fill in:
   - Examination Details (e.g. "Fourth Semester FYUGP Examination")
   - Programme / Major (e.g. "Physics")
   - Month & Year (e.g. "May 2026")
3. Upload the class mark-list PDF (one student's mark statement per
   page).
4. Click **Generate Report**.
5. Scroll through the on-screen report, then:
   - Click **"Download Full Report (.docx)"** for the formatted Word
     document (logo letterhead + all 9 sections + charts + signature
     blocks), or
   - Click **"Download Student Table Only (CSV)"** for just the raw
     consolidated table.

## Notes on the numbers shown

- **Total Credits**, **SGPA**, and **Grade Awarded** are read directly
  from the official mark statement. If the University's own sheet shows
  `--` for these (which it does whenever the result is a Fail), the app
  shows `--` as well, to stay faithful to the source document.
- A course is counted as "Failed" for a student when that course's
  Grade column reads **F**.
- In the **Grade Distribution** and **Grade-wise List of Students**
  sections, a student's overall grade is their printed "Grade Awarded"
  if they passed, or **"F"** if their overall result is Fail (since the
  official sheet prints `--`, not a letter grade, for a failed result).
  These two sections are always ordered **O, A+, A, B+, B, C, P, F**.
- **Average / Highest / Lowest SGPA** are computed only over students
  with a valid (numeric) SGPA, i.e. students who passed.

## Compatibility

The parser is built for the standard University of Kerala FYUGP
"Mark cum Grade Statement" PDF layout (one page per student, with a
course table containing Course Code / Title / Credits / CCA / ESE /
Course Total / GP / CP / Grade columns, followed by a `Total` row and a
`Result: ... SGPA: ... Grade awarded: ...` line). If your college uses a
slightly different template, minor tweaks to `parse_marklist_pdf()` in
`app.py` may be needed — the column layout is clearly commented in the
code.
