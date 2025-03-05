import re
import pandas as pd
from io import BytesIO
import streamlit as st
from PyPDF2 import PdfReader

# ---------- Backend Functions ----------

def clean_mark(token):
    """
    If token contains '/', return only the part before the slash.
    Otherwise, return the token as is.
    """
    if "/" in token:
        return token.split("/")[0]
    return token

def extract_text_from_pdf(pdf_bytes):
    """
    Extract text from a PDF file given as bytes.
    """
    reader = PdfReader(BytesIO(pdf_bytes))
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def auto_detect_subjects(text):
    """
    Auto-detect subject base names from lines that start with a course code
    and contain a '*' token. Returns a list of base subject names.
    """
    subjects = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if re.match(r"^\d", line) and "*" in line:
            tokens = line.split()
            try:
                star_index = tokens.index("*")
            except ValueError:
                continue
            subject = " ".join(tokens[1:star_index]).strip()
            if subject and subject not in subjects:
                subjects.append(subject)
    return subjects

def parse_student_file_from_text(text):
    """
    Parse the extracted text from the PDF and return a list of student dictionaries.
    Each student record is built dynamically – course lines add keys for subjects,
    and extra summary lines (containing SGPA, TOTAL CREDITS, CGPA, etc.) are parsed.
    """
    lines = text.splitlines()
    students = []
    current_student = None

    def new_student():
        # Start with basic fields. We'll later add extra summary fields dynamically.
        return {
            "Seat No.": "-",
            "Name of Student": "-",
            # We use these keys as our common fields.
            "SGPA": "-",
            "TOTAL CREDITS EARNED": "-",
            "CGPA": "-",
            "Total": "",
            "%": "",
            "Remarks": ""
        }

    # Regex to match a student header line.
    header_regex = re.compile(
        r"SEAT NO\.\:\s*(\S+)\s*NAME\s*:\s*(.*?)\s*MOTHER\s*:\s*.*?PRN\s*:\s*\d+.*?CLG\.\:\s*(\S+)"
    )
    
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Skip extra header/footer lines.
        if (line.startswith("COURSE NAME") or line.startswith("SEM.:") or 
            line.startswith("............") or line.startswith("PAGE :-") or
            line.startswith("COLLEGE:") or line.startswith("BRANCH CODE")):
            continue

        # If we find a header line for a new student:
        header_match = header_regex.search(line)
        if header_match:
            if current_student is not None:
                # Before appending, if CGPA not set from extra info, set it equal to SGPA.
                if current_student.get("CGPA", "-") == "-":
                    current_student["CGPA"] = current_student.get("SGPA", "-")
                students.append(current_student)
            current_student = new_student()
            current_student["Seat No."] = header_match.group(1)
            current_student["Name of Student"] = header_match.group(2).strip()
            continue

        # Process extra summary lines – lines that contain "SGPA" and ":".
        if "SGPA" in line and ":" in line:
            # Find all key-value pairs for SGPA type fields.
            pairs = re.findall(r"([A-Z\s/]+SGPA)\s*:\s*([\d.]+)", line)
            for key, value in pairs:
                current_student[key.strip()] = value.strip()
            # Look for TOTAL CREDITS EARNED.
            tc_match = re.search(r"TOTAL CREDITS EARNED\s*:\s*(\d+)", line)
            if tc_match:
                current_student["TOTAL CREDITS EARNED"] = tc_match.group(1)
            # Look for TOTAL GRADE POINTS / TOTAL CREDITS.
            tg_match = re.search(r"TOTAL GRADE POINTS\s*/\s*TOTAL CREDITS\s*:\s*([\d/]+)", line)
            if tg_match:
                current_student["TOTAL GRADE POINTS/TOTAL CREDITS"] = tg_match.group(1)
            # Look for CGPA.
            cg_match = re.search(r"CGPA\s*:\s*([\d.]+)", line)
            if cg_match:
                current_student["CGPA"] = cg_match.group(1)
            # If the line contains a remark (e.g. "FIRST CLASS WITH DISTINCTION"), store it.
            if "CLASS" in line.upper() or "DISTINCTION" in line.upper():
                current_student["Remarks"] = line
            continue

        # Process course lines: they start with a course code (digit) and contain a '*' token.
        if re.match(r"^\d", line) and "*" in line:
            tokens = line.split()
            try:
                star_index = tokens.index("*")
            except ValueError:
                continue
            base_subject = " ".join(tokens[1:star_index]).strip()
            marks = tokens[star_index+1:]
            marks = [clean_mark(tok) for tok in marks]
            
            # Determine the subject type based on the base_subject text.
            if "LABORATORY PRACTICE" in base_subject.upper():
                if "III" in base_subject.upper():
                    if len(marks) >= 5:
                        key_tw = base_subject + " (TW)"
                        key_pr = base_subject + " (PR)"
                        current_student[key_tw] = marks[3]
                        current_student[key_pr] = marks[4]
                elif "IV" in base_subject.upper():
                    if len(marks) >= 4:
                        key_tw = base_subject + " (TW)"
                        current_student[key_tw] = marks[3]
                else:
                    if len(marks) >= 5:
                        key_tw = base_subject + " (TW)"
                        key_pr = base_subject + " (PR)"
                        current_student[key_tw] = marks[3]
                        current_student[key_pr] = marks[4]
            elif "MOOC" in base_subject.upper():
                if len(marks) >= 4:
                    key_mark = base_subject + " (Mark)"
                    current_student[key_mark] = marks[3]
            else:
                # Assume theory subject: expect at least 3 tokens.
                if len(marks) >= 3:
                    key_insem = base_subject + " (Insem)"
                    key_ese = base_subject + " (ESE)"
                    key_total = base_subject + " (Total)"
                    current_student[key_insem] = marks[0]
                    current_student[key_ese] = marks[1]
                    current_student[key_total] = marks[2]
    if current_student is not None:
        if current_student.get("CGPA", "-") == "-":
            current_student["CGPA"] = current_student.get("SGPA", "-")
        students.append(current_student)
    return students

def create_excel_in_memory(students, selected_subjects_str):
    """
    Create an Excel file (as bytes) from the student data.
    The selected_subjects_str is a comma-separated list of base subject names.
    Only keys (columns) matching these subjects (or extra summary fields) are shown.
    """
    # Define common columns to always include.
    common_cols = ["Seat No.", "Name of Student", "SGPA", "TOTAL CREDITS EARNED", "CGPA", "Total", "%", "Remarks", "TOTAL GRADE POINTS/TOTAL CREDITS"]
    
    # Gather all keys from student records.
    dynamic_keys = set()
    for s in students:
        for k in s.keys():
            if k not in common_cols:
                dynamic_keys.add(k)
    
    # Process user input for subject base names.
    user_subjects = [x.strip() for x in selected_subjects_str.split(",") if x.strip()]
    final_subject_cols = []
    for subj in user_subjects:
        for key in sorted(dynamic_keys):
            if key.upper().startswith(subj.upper()):
                if key not in final_subject_cols:
                    final_subject_cols.append(key)
    # Optionally add any remaining dynamic keys.
    for key in sorted(dynamic_keys):
        if key not in final_subject_cols:
            final_subject_cols.append(key)
    
    # Final columns: serial, common fields, then dynamic subject columns.
    final_cols = ["Sr.", "Seat No.", "Name of Student"] + final_subject_cols + common_cols[2:]  # excluding Seat No. and Name (already included)
    
    rows = []
    sr = 1
    for s in students:
        row = {"Sr.": sr}
        for col in final_cols[1:]:
            row[col] = s.get(col, "-")
        rows.append(row)
        sr += 1
    df = pd.DataFrame(rows, columns=final_cols)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Student Data")
    return output.getvalue()

# ---------- Multi-Page Streamlit App ----------

# Sidebar for navigation.
page = st.sidebar.selectbox("Navigation", ["Home", "Contact"])

if page == "Home":
    st.set_page_config(page_title="Result Ledger Parser", layout="wide")
    st.title("Dynamic Result Ledger Parser from PDF")
    
    st.markdown("""
    **Instructions:**
    
    1. **Upload a PDF file:**  
       The file should contain result ledger data formatted as follows:
       - A header line starting with **"SEAT NO.:"** followed by student details.
       - Course lines with subject marks (e.g.,  
         `410241 DESIGN & ANALYSIS OF ALGO. * 025/030 053/070 078/100 ...`).
       - One or more lines containing various SGPA values, TOTAL CREDITS, CGPA, etc.  
         (For example:  
         `FOURTH YEAR SGPA : 7.53, TOTAL CREDITS EARNED : 40`  
         `FE SGPA : 8.61 SE SGPA : 7.59 TE SGPA : 7.64`  
         `TOTAL GRADE POINTS / TOTAL CREDITS : 1335/170 CGPA : 7.85 FIRST CLASS WITH DISTINCTION`)
    
    2. **Auto-Detection of Subjects:**  
       The app will auto-detect subject base names from the PDF (the text between the course code and the "*" marker).  
       You can review and modify these names (comma separated) below.  
       For example: **DESIGN & ANALYSIS OF ALGO., MACHINE LEARNING, BLOCKCHAIN TECHNOLOGY, OBJ. ORIENTED MODL. & DESG., SOFT. TEST. & QLTY ASSURANCE, LABORATORY PRACTICE - III, LABORATORY PRACTICE - IV, PROJECT STAGE - I, MOOC - LEARN NEW SKILLS**.
    
    3. **Generate the Ledger:**  
       Confirm the subject names by checking the box, then click **"Generate Excel"** to process the data and download the resulting Excel file.
    """)
    
    uploaded_pdf = st.file_uploader("Choose a PDF file", type=["pdf"])
    
    if uploaded_pdf is not None:
        pdf_bytes = uploaded_pdf.getvalue()
        extracted_text = extract_text_from_pdf(pdf_bytes)
        
        # Auto-detect subject base names.
        detected_subjects = auto_detect_subjects(extracted_text)
        if detected_subjects:
            detected_str = ", ".join(detected_subjects)
            st.info(f"Auto-detected subject base names: **{detected_str}**")
        else:
            st.warning("No subjects were automatically detected. Please enter subject base names manually.")
        
        default_subject_input = detected_str if detected_subjects else ""
        subject_names_input = st.text_input("Enter subject base names (comma separated)", default_subject_input)
        
        proceed = st.checkbox("Proceed with these subjects?")
        
        if proceed:
            if st.button("Generate Excel"):
                students = parse_student_file_from_text(extracted_text)
                excel_bytes = create_excel_in_memory(students, subject_names_input)
                st.download_button(
                    label="Download Excel File",
                    data=excel_bytes,
                    file_name="output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

elif page == "Contact":
    st.set_page_config(page_title="Contact", layout="centered")
    st.title("Contact Information")
    st.markdown("""
    **For Support:**
    
    **Name:** Yash Bari  
    **Email:** [yashbari99@gmail.com](mailto:yashbari99@gmail.com)  
    **Phone:** 9834148536
    
    If you have any questions or need assistance with this application, please feel free to reach out.
    """)
