from flask import Flask, render_template, request, redirect, session, send_file
import sqlite3
import os
import re
import hashlib

from pypdf import PdfReader
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT


app = Flask(__name__)

app.secret_key = "resume_screening_secret_key"

UPLOAD_FOLDER = "uploads"
WORD_FILE = "candidate_details.docx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ==========================================================
# DATABASE
# ==========================================================

def create_database():

    conn = sqlite3.connect("resumes.db")
    cursor = conn.cursor()

    # ======================================================
    # REQUIRED SKILLS TABLE
    # ======================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS required_skills (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            skill TEXT UNIQUE NOT NULL
        )
    """)

    # ======================================================
    # CANDIDATES TABLE
    # ======================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS candidates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidate_name TEXT,
            email TEXT,
            phone TEXT,
            education TEXT,
            skills TEXT,
            required_skills TEXT,
            matched_skills TEXT,
            missing_skills TEXT,
            match_percentage INTEGER,
            resume_filename TEXT,
            candidate_key TEXT UNIQUE
        )
    """)

    # ======================================================
    # UPLOAD COUNT TABLE
    # ======================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS upload_count (
            id INTEGER PRIMARY KEY,
            total_uploads INTEGER
        )
    """)

    # ======================================================
    # CREATE FIRST UPLOAD COUNT RECORD
    # ======================================================

    cursor.execute("""
        SELECT id
        FROM upload_count
        WHERE id = 1
    """)

    count = cursor.fetchone()

    if count is None:

        cursor.execute("""
            INSERT INTO upload_count
            (id, total_uploads)
            VALUES (1, 0)
        """)

    conn.commit()
    conn.close()


# ==========================================================
# LOGIN CHECK
# ==========================================================

def check_login():

    return session.get("logged_in") is True


# ==========================================================
# HOME / DASHBOARD
# ==========================================================

@app.route("/")
def home():

    if not check_login():
        return redirect("/login")

    conn = sqlite3.connect("resumes.db")
    cursor = conn.cursor()

    # ======================================================
    # TOTAL RESUMES UPLOADED
    # ======================================================

    cursor.execute("""
        SELECT total_uploads
        FROM upload_count
        WHERE id = 1
    """)

    result = cursor.fetchone()

    if result:
        total_uploads = result[0]
    else:
        total_uploads = 0

    # ======================================================
    # COUNT ALL UNIQUE CANDIDATES
    # ======================================================

    cursor.execute("""
        SELECT COUNT(*)
        FROM candidates
    """)

    unique_candidates = cursor.fetchone()[0]

    conn.close()

    return render_template(
        "index.html",
        total_uploads=total_uploads,
        unique_candidates=unique_candidates
    )


# ==========================================================
# LOGIN
# ==========================================================

@app.route("/login", methods=["GET", "POST"])
def login():

    error = ""

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        ).strip()

        if username == "admin" and password == "admin123":

            session["logged_in"] = True

            return redirect("/")

        else:

            error = "Invalid username or password"

    return render_template(
        "login.html",
        error=error
    )


# ==========================================================
# LOGOUT
# ==========================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect("/login")


# ==========================================================
# REQUIRED SKILLS
# ==========================================================

@app.route("/skills", methods=["GET", "POST"])
def skills():

    if not check_login():
        return redirect("/login")

    conn = sqlite3.connect("resumes.db")
    cursor = conn.cursor()

    if request.method == "POST":

        skill = request.form.get(
            "skill",
            ""
        ).strip()

        if skill:

            try:

                cursor.execute(
                    """
                    INSERT INTO required_skills
                    (skill)
                    VALUES (?)
                    """,
                    (skill,)
                )

                conn.commit()

            except sqlite3.IntegrityError:

                pass

    cursor.execute("""
        SELECT id, skill
        FROM required_skills
        ORDER BY id
    """)

    required_skills = cursor.fetchall()

    conn.close()

    return render_template(
        "skills.html",
        required_skills=required_skills
    )


# ==========================================================
# DELETE ONE SKILL
# ==========================================================

@app.route("/delete_skill/<int:skill_id>")
def delete_skill(skill_id):

    if not check_login():
        return redirect("/login")

    conn = sqlite3.connect("resumes.db")
    cursor = conn.cursor()

    cursor.execute(
        """
        DELETE FROM required_skills
        WHERE id = ?
        """,
        (skill_id,)
    )

    conn.commit()
    conn.close()

    return redirect("/skills")


# ==========================================================
# DELETE ALL SKILLS
# ==========================================================

@app.route("/delete_all_skills")
def delete_all_skills():

    if not check_login():
        return redirect("/login")

    conn = sqlite3.connect("resumes.db")
    cursor = conn.cursor()

    cursor.execute(
        "DELETE FROM required_skills"
    )

    conn.commit()
    conn.close()

    return redirect("/skills")


# ==========================================================
# PDF TEXT EXTRACTION
# ==========================================================

def extract_pdf_text(file_path):

    text = ""

    reader = PdfReader(file_path)

    for page in reader.pages:

        page_text = page.extract_text()

        if page_text:

            text += page_text + "\n"

    return text


# ==========================================================
# DOCX TEXT EXTRACTION
# ==========================================================

def extract_docx_text(file_path):

    document = Document(file_path)

    text = ""

    for paragraph in document.paragraphs:

        text += paragraph.text + "\n"

    return text


# ==========================================================
# EMAIL EXTRACTION
# ==========================================================

def extract_email(text):

    pattern = (
        r'\b[A-Za-z0-9._%+-]+'
        r'@[A-Za-z0-9.-]+'
        r'\.[A-Za-z]{2,}\b'
    )

    match = re.search(
        pattern,
        text
    )

    if match:

        return match.group()

    return "Not Found"


# ==========================================================
# PHONE EXTRACTION
# ==========================================================

def extract_phone(text):

    pattern = (
        r'(?<!\d)'
        r'(?:\+91[\s-]?)?'
        r'[6-9]\d{9}'
        r'(?!\d)'
    )

    match = re.search(
        pattern,
        text
    )

    if match:

        return match.group()

    return "Not Found"


# ==========================================================
# NAME EXTRACTION
# ==========================================================

def extract_name(text):

    lines = text.split("\n")

    for line in lines:

        line = line.strip()

        if not line:
            continue

        lower_line = line.lower()

        # If resume contains Name:
        if "name:" in lower_line:

            name = line.split(
                ":",
                1
            )[1].strip()

            if name:

                return name

        if "email" in lower_line:
            continue

        if "phone" in lower_line:
            continue

        if "mobile" in lower_line:
            continue

        if "resume" in lower_line:
            continue

        words = line.split()

        if 2 <= len(words) <= 4:

            valid = True

            for word in words:

                if not word.replace(
                    ".",
                    ""
                ).isalpha():

                    valid = False
                    break

            if valid:

                return line

    return "Not Found"


# ==========================================================
# EDUCATION EXTRACTION
# ==========================================================

def extract_education(text):

    education_keywords = [

        "B.Tech",
        "B.E",
        "B.Sc",
        "BCA",
        "M.Tech",
        "MCA",
        "M.Sc",
        "MBA",
        "Bachelor",
        "Master",
        "Degree",
        "Engineering"

    ]

    found = []

    text_lower = text.lower()

    for keyword in education_keywords:

        if keyword.lower() in text_lower:

            if keyword not in found:

                found.append(keyword)

    if found:

        return ", ".join(found)

    return "Not Found"


# ==========================================================
# RESUME SKILLS
# ==========================================================

def extract_skills(text):

    possible_skills = [

        "Python",
        "SQL",
        "Machine Learning",
        "Deep Learning",
        "Django",
        "Flask",
        "JavaScript",
        "Java",
        "HTML",
        "CSS",
        "React.js",
        "React",
        "Node.js",
        "MongoDB",
        "MySQL",
        "Power BI",
        "Tableau",
        "Excel",
        "AWS",
        "Azure",
        "Databricks",
        "PySpark",
        "Spark",
        "Pandas",
        "NumPy",
        "Scikit-learn",
        "NLP",
        "Artificial Intelligence",
        "C++",
        "C",
        "Git",
        "GitHub"

    ]

    found = []

    text_lower = text.lower()

    for skill in possible_skills:

        if skill.lower() in text_lower:

            if skill not in found:

                found.append(skill)

    return list(
        dict.fromkeys(found)
    )


# ==========================================================
# CREATE UNIQUE CANDIDATE KEY
# ==========================================================

def create_candidate_key(
    email,
    phone,
    candidate_name,
    resume_text
):

    if email != "Not Found":

        return (
            "email:"
            + email.lower().strip()
        )

    if phone != "Not Found":

        return (
            "phone:"
            + phone.strip()
        )

    unique_text = (

        candidate_name.lower().strip()

        + "|"

        + resume_text.lower().strip()

    )

    hash_value = hashlib.sha256(
        unique_text.encode("utf-8")
    ).hexdigest()

    return (
        "resume:"
        + hash_value
    )


# ==========================================================
# UPDATE WORD FILE
# ==========================================================

def update_word_file():

    document = Document()

    # ======================================================
    # LANDSCAPE PAGE
    # ======================================================

    section = document.sections[0]

    section.orientation = WD_ORIENT.LANDSCAPE

    section.page_width = Inches(11)

    section.page_height = Inches(8.5)

    section.top_margin = Inches(0.4)

    section.bottom_margin = Inches(0.4)

    section.left_margin = Inches(0.4)

    section.right_margin = Inches(0.4)

    # ======================================================
    # TITLE
    # ======================================================

    document.add_heading(
        "RESUME SCREENING SYSTEM",
        level=1
    )

    document.add_paragraph(
        "All Candidates with Manually Entered Match Percentage"
    )

    conn = sqlite3.connect(
        "resumes.db"
    )

    cursor = conn.cursor()

    # ======================================================
    # TOTAL UPLOADS
    # ======================================================

    cursor.execute("""
        SELECT total_uploads
        FROM upload_count
        WHERE id = 1
    """)

    result = cursor.fetchone()

    if result:

        total_uploads = result[0]

    else:

        total_uploads = 0

    # ======================================================
    # TOTAL UNIQUE CANDIDATES
    # ======================================================

    cursor.execute("""
        SELECT COUNT(*)
        FROM candidates
    """)

    unique_candidates = cursor.fetchone()[0]

    document.add_paragraph(
        "Total Resumes Uploaded: "
        + str(total_uploads)
    )

    document.add_paragraph(
        "Total Unique Candidates: "
        + str(unique_candidates)
    )

    document.add_paragraph("")

    # ======================================================
    # GET ALL CANDIDATES
    # ======================================================

    cursor.execute("""
        SELECT
            candidate_name,
            email,
            phone,
            education,
            matched_skills,
            missing_skills,
            match_percentage,
            resume_filename
        FROM candidates
        ORDER BY id DESC
    """)

    candidates = cursor.fetchall()

    # ======================================================
    # HORIZONTAL TABLE
    # ======================================================

    if candidates:

        table = document.add_table(
            rows=1,
            cols=8
        )

        table.style = "Table Grid"

        header = table.rows[0].cells

        header[0].text = "Name"
        header[1].text = "Email"
        header[2].text = "Phone"
        header[3].text = "Education"
        header[4].text = "Matched Skills"
        header[5].text = "Missing Skills"
        header[6].text = "Match %"
        header[7].text = "Resume"

        # ==================================================
        # COLUMN WIDTHS
        # ==================================================

        widths = [
            1.2,
            1.7,
            1.2,
            1.1,
            2.0,
            1.5,
            0.7,
            1.5
        ]

        for candidate in candidates:

            row = table.add_row().cells

            row[0].text = str(candidate[0])
            row[1].text = str(candidate[1])
            row[2].text = str(candidate[2])
            row[3].text = str(candidate[3])
            row[4].text = str(candidate[4])
            row[5].text = str(candidate[5])
            row[6].text = str(candidate[6]) + "%"
            row[7].text = str(candidate[7])

        # ==================================================
        # SET WIDTHS AND FONT SIZE
        # ==================================================

        for row in table.rows:

            for index, width in enumerate(widths):

                row.cells[index].width = Inches(width)

                for paragraph in row.cells[index].paragraphs:

                    for run in paragraph.runs:

                        run.font.size = Pt(8)

    else:

        document.add_paragraph(
            "No candidates have been uploaded yet."
        )

    conn.close()

    # ======================================================
    # SAVE WORD FILE
    # ======================================================

    document.save(
        WORD_FILE
    )


# ==========================================================
# UPLOAD RESUME
# ==========================================================

@app.route("/upload", methods=["POST"])
def upload():

    if not check_login():
        return redirect("/login")

    # ======================================================
    # GET MANUAL PERCENTAGE
    # ======================================================

    manual_percentage = request.form.get(
        "match_percentage",
        ""
    ).strip()

    if manual_percentage == "":

        return "Please enter match percentage"

    try:

        match_percentage = int(
            manual_percentage
        )

    except ValueError:

        return "Please enter a valid percentage"

    # ======================================================
    # CHECK PERCENTAGE RANGE
    # ======================================================

    if match_percentage < 0 or match_percentage > 100:

        return "Percentage must be between 0 and 100"

    # ======================================================
    # GET RESUME
    # ======================================================

    resume = request.files.get(
        "resume"
    )

    if resume is None:

        return "Please select a resume"

    if resume.filename == "":

        return "Please select a resume"

    filename = resume.filename

    extension = os.path.splitext(
        filename
    )[1].lower()

    # ======================================================
    # CHECK FILE TYPE
    # ======================================================

    if extension not in [
        ".pdf",
        ".docx"
    ]:

        return (
            "Only PDF and DOCX "
            "files are supported"
        )

    # ======================================================
    # SAVE RESUME
    # ======================================================

    file_path = os.path.join(
        UPLOAD_FOLDER,
        filename
    )

    resume.save(file_path)

    # ======================================================
    # EXTRACT TEXT
    # ======================================================

    try:

        if extension == ".pdf":

            resume_text = extract_pdf_text(
                file_path
            )

        else:

            resume_text = extract_docx_text(
                file_path
            )

    except Exception as e:

        return (
            "Error reading resume: "
            + str(e)
        )

    # ======================================================
    # EXTRACT CANDIDATE DETAILS
    # ======================================================

    candidate_name = extract_name(
        resume_text
    )

    email = extract_email(
        resume_text
    )

    phone = extract_phone(
        resume_text
    )

    education = extract_education(
        resume_text
    )

    # ======================================================
    # GET REQUIRED SKILLS
    # ======================================================

    conn = sqlite3.connect(
        "resumes.db"
    )

    cursor = conn.cursor()

    cursor.execute("""
        SELECT skill
        FROM required_skills
        ORDER BY id
    """)

    required_skills = [

        row[0]

        for row in cursor.fetchall()

    ]

    # ======================================================
    # MATCH SKILLS
    # ======================================================

    matched_skills = []

    missing_skills = []

    resume_text_lower = resume_text.lower()

    for skill in required_skills:

        if skill.lower() in resume_text_lower:

            if skill not in matched_skills:

                matched_skills.append(
                    skill
                )

        else:

            if skill not in missing_skills:

                missing_skills.append(
                    skill
                )

    # ======================================================
    # REMOVE DUPLICATE MATCHED SKILLS
    # ======================================================

    matched_skills = list(
        dict.fromkeys(
            matched_skills
        )
    )

    # ======================================================
    # REMOVE DUPLICATE MISSING SKILLS
    # ======================================================

    missing_skills = list(
        dict.fromkeys(
            missing_skills
        )
    )

    # ======================================================
    # CONVERT SKILLS TO TEXT
    # ======================================================

    matched_skills_text = ", ".join(
        matched_skills
    )

    required_skills_text = ", ".join(
        required_skills
    )

    missing_skills_text = ", ".join(
        missing_skills
    )

    # ======================================================
    # CREATE UNIQUE CANDIDATE KEY
    # ======================================================

    candidate_key = create_candidate_key(

        email,

        phone,

        candidate_name,

        resume_text

    )

    # ======================================================
    # CHECK EXISTING CANDIDATE
    # ======================================================

    cursor.execute("""
        SELECT id
        FROM candidates
        WHERE candidate_key = ?
    """, (candidate_key,))

    existing_candidate = cursor.fetchone()

    # ======================================================
    # UPDATE EXISTING CANDIDATE
    # ======================================================

    if existing_candidate:

        cursor.execute("""
            UPDATE candidates
            SET
                candidate_name = ?,
                email = ?,
                phone = ?,
                education = ?,
                skills = ?,
                required_skills = ?,
                matched_skills = ?,
                missing_skills = ?,
                match_percentage = ?,
                resume_filename = ?
            WHERE candidate_key = ?
        """, (

            candidate_name,

            email,

            phone,

            education,

            # Only matched skills
            matched_skills_text,

            required_skills_text,

            matched_skills_text,

            missing_skills_text,

            # Exact manual percentage
            match_percentage,

            filename,

            candidate_key

        ))

    # ======================================================
    # INSERT NEW CANDIDATE
    # ======================================================

    else:

        cursor.execute("""
            INSERT INTO candidates
            (
                candidate_name,
                email,
                phone,
                education,
                skills,
                required_skills,
                matched_skills,
                missing_skills,
                match_percentage,
                resume_filename,
                candidate_key
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (

            candidate_name,

            email,

            phone,

            education,

            # Only matched skills
            matched_skills_text,

            required_skills_text,

            matched_skills_text,

            missing_skills_text,

            # Exact manual percentage
            match_percentage,

            filename,

            candidate_key

        ))

    # ======================================================
    # INCREASE TOTAL UPLOAD COUNT
    # ======================================================

    cursor.execute("""
        UPDATE upload_count
        SET total_uploads =
            total_uploads + 1
        WHERE id = 1
    """)

    conn.commit()

    conn.close()

    # ======================================================
    # UPDATE WORD FILE
    # ======================================================

    update_word_file()

    # ======================================================
    # SHOW RESULT
    # ======================================================

    return render_template(

        "result.html",

        candidate_name=candidate_name,

        email=email,

        phone=phone,

        education=education,

        resume_skills=matched_skills,

        required_skills=required_skills,

        matched_skills=matched_skills,

        missing_skills=missing_skills,

        # Exact manually entered percentage
        match_percentage=match_percentage,

        filename=filename

    )


# ==========================================================
# CANDIDATES PAGE
# ==========================================================

@app.route("/candidates")
def candidates():

    if not check_login():

        return redirect("/login")

    conn = sqlite3.connect(
        "resumes.db"
    )

    cursor = conn.cursor()

    # ======================================================
    # SHOW ALL CANDIDATES
    # ======================================================

    cursor.execute("""
        SELECT
            id,
            candidate_name,
            email,
            phone,
            education,
            skills,
            required_skills,
            matched_skills,
            missing_skills,
            match_percentage,
            resume_filename
        FROM candidates
        ORDER BY id DESC
    """)

    candidates = cursor.fetchall()

    conn.close()

    return render_template(
        "candidates.html",
        candidates=candidates
    )


# ==========================================================
# DOWNLOAD WORD REPORT
# ==========================================================

@app.route("/download_word")
def download_word():

    if not check_login():

        return redirect("/login")

    # ======================================================
    # CREATE LATEST WORD FILE
    # ======================================================

    update_word_file()

    return send_file(

        WORD_FILE,

        as_attachment=True,

        download_name="candidate_details.docx"

    )


# ==========================================================
# RUN APPLICATION
# ==========================================================

if __name__ == "__main__":

    create_database()

    app.run(
        debug=True
    )