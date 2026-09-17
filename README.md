# Result Analysis Report Generator

A Streamlit app that takes a University "Mark cum Grade Statement" PDF
(one page per student — the standard University of Kerala FYUGP layout)
and automatically generates a full **Result Analysis Report**:

- Consolidated table: Sl.No, Name, APAAR ID, Total Marks, Total Credits,
  SGPA, Grade Awarded, Status
- Appeared / Passed / Failed summary + Pass % + bar chart
- Top-3 rank holders (all tied students included)
- Subject-wise failure counts (course-code wise) + bar chart, and the
  subject with the highest number of failures is called out
- Course code → full course name reference table
- Extra analysis: grade distribution, average/highest/lowest SGPA

Editable fields at the top of the report: **College/Institution name**,
**Examination details** (e.g. "Fourth Semester FYUGP Examination"),
**Programme/Major** (e.g. "Physics"), and **Month & Year** (e.g. "May 2026").

---

## 1. Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).

## 2. Deploy for free on Streamlit Community Cloud (via GitHub)

1. Create a new GitHub repository (public or private) and push these two
   files to it:
   - `app.py`
   - `requirements.txt`

   ```bash
   git init
   git add app.py requirements.txt README.md
   git commit -m "Result analysis app"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```

2. Go to **https://share.streamlit.io** and sign in with your GitHub
   account.

3. Click **"New app"**, choose your repository and branch (`main`), and
   set the **Main file path** to `app.py`.

4. Click **Deploy**. Streamlit Cloud will install everything from
   `requirements.txt` automatically and give you a public URL such as
   `https://<your-app-name>.streamlit.app`.

5. Any time you push new commits to the repo, the deployed app
   auto-updates.

## 3. How to use the app

1. Open the app.
2. In the left sidebar, fill in:
   - College / Institution Name
   - Examination Details (e.g. "Fourth Semester FYUGP Examination")
   - Programme / Major (e.g. "Physics")
   - Month & Year (e.g. "May 2026")
3. Upload the class mark-list PDF (the one with one student's mark
   statement per page).
4. Click **Generate Report**.
5. Scroll through the report; use the **Download CSV** button to export
   the consolidated student table if needed.

## Notes on the numbers shown

- **Total Marks** is *computed* by the app as the sum of each course's
  "Course Total (CCA+ESE)" value across all courses on the student's
  mark statement. The official University sheet does not print this
  overall figure directly.
- **Total Credits**, **SGPA**, and **Grade Awarded** are read directly
  from the official mark statement. If the University's own sheet shows
  `--` for these (which it does whenever the result is a Fail), the app
  shows `--` as well, to stay faithful to the source document.
- A course is counted as "Failed" for a student when that course's
  Grade column reads **F**.

## Compatibility

The parser is built for the standard University of Kerala FYUGP
"Mark cum Grade Statement" PDF layout (one page per student, with a
course table containing Course Code / Title / Credits / CCA / ESE /
Course Total / GP / CP / Grade columns, followed by a `Total` row and a
`Result: ... SGPA: ... Grade awarded: ...` line). If your college uses a
slightly different template, minor tweaks to `parse_marklist_pdf()` in
`app.py` may be needed — the column layout is clearly commented in the
code.
