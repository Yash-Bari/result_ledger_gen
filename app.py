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
    Auto-detect subject base names from lines that start with a course code and contain a '*' token.
    Returns a list of base subject names.
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
    Student records are built dynamically by adding keys for each subject encountered.
    """
    lines = text.splitlines()
    students = []
    current_student = None

    def new_student():
        return {
            "Seat No.": "-",
            "Name of Student": "-",
            "Mother's Name": "-",
            "PRN": "-",
            "College Code": "-",
            "SGPA": "-",
            "Total Credits": "-",
            "CGPA": "-",
            "Total": "",
            "%": ""
        }

    header_regex = re.compile(
        r"SEAT NO\.\:\s*(\S+)\s*NAME\s*:\s*(.*?)\s*MOTHER\s*:\s*(.*?)\s*PRN\s*:\s*(\S+)\s*CLG\.\:\s*(\S+)"
    )
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if (line.startswith("COURSE NAME") or line.startswith("SEM.:") or 
            line.startswith("............") or line.startswith("PAGE :-") or
            line.startswith("COLLEGE:") or line.startswith("BRANCH CODE")):
            continue

        header_match = header_regex.search(line)
        if header_match:
            if current_student is not None:
                current_student["CGPA"] = current_student["SGPA"]
                
                # Calculate total marks and percentage
                total_marks = 0
                max_marks = 0
                for key, value in current_student.items():
                    if key.endswith(" (Total)") and value not in ["-", "AB", "FF"]:
                        try:
                            total_marks += int(value)
                            max_marks += 100  # Assuming each subject is out of 100
                        except ValueError:
                            pass
                    elif (key.endswith(" (TW)") or key.endswith(" (PR)")) and value not in ["-", "AB", "FF"]:
                        try:
                            total_marks += int(value)
                            max_marks += 50  # Assuming TW/PR are out of 50
                        except ValueError:
                            pass
                
                if max_marks > 0:
                    current_student["Total"] = str(total_marks)
                    current_student["%"] = str(round((total_marks / max_marks) * 100, 2))
                
                students.append(current_student)
                
            current_student = new_student()
            current_student["Seat No."] = header_match.group(1)
            current_student["Name of Student"] = header_match.group(2).strip()
            current_student["Mother's Name"] = header_match.group(3).strip()
            current_student["PRN"] = header_match.group(4).strip()
            current_student["College Code"] = header_match.group(5).strip()
            continue

        if line.startswith("SGPA1 :"):
            sgpa_match = re.search(r"SGPA1\s*:\s*([\d.]+|--)", line)
            tc_match = re.search(r"TOTAL CREDITS EARNED\s*:\s*(\d+)", line)
            if sgpa_match:
                current_student["SGPA"] = sgpa_match.group(1)
            if tc_match:
                current_student["Total Credits"] = tc_match.group(1)
            continue

        if re.match(r"^\d", line) and "*" in line:
            tokens = line.split()
            try:
                star_index = tokens.index("*")
            except ValueError:
                continue
                
            course_code = tokens[0]
            base_subject = " ".join(tokens[1:star_index]).strip()
            marks = tokens[star_index+1:]
            
            # Handle special cases like "AB", "FF", etc.
            processed_marks = []
            for mark in marks:
                if mark in ["AB", "AB/070", "AB/050", "FF", "PP"]:
                    processed_marks.append(mark.split("/")[0] if "/" in mark else mark)
                else:
                    processed_marks.append(clean_mark(mark))
            
            # Store course code as a separate field
            current_student[base_subject + " (Code)"] = course_code
            
            if "LABORATORY PRACTICE" in base_subject.upper() or "PROJECT" in base_subject.upper():
                # Fix for Lab Practice and Project subjects
                tw_index = -1
                pr_index = -1
                
                for i, token in enumerate(marks):
                    if "TW" in token or token.endswith("/050"):
                        tw_index = i
                    elif "PR" in token or (i > tw_index and token.endswith("/050")):
                        pr_index = i
                
                if tw_index >= 0 and tw_index < len(processed_marks):
                    key_tw = base_subject + " (TW)"
                    current_student[key_tw] = processed_marks[tw_index]
                
                if pr_index >= 0 and pr_index < len(processed_marks):
                    key_pr = base_subject + " (PR)"
                    current_student[key_pr] = processed_marks[pr_index]
            elif "MOOC" in base_subject.upper():
                # Handle MOOC subjects
                if len(processed_marks) >= 1:
                    for i, mark in enumerate(processed_marks):
                        if mark in ["PP", "FF"]:
                            key_mark = base_subject + " (Status)"
                            current_student[key_mark] = mark
            else:
                # Regular subjects with ISE, ESE, and Total
                ise_index = -1
                ese_index = -1
                total_index = -1
                
                for i, token in enumerate(marks):
                    if i == 0 and ("ISE" in token or token.endswith("/030")):
                        ise_index = i
                    elif i == 1 and ("ESE" in token or token.endswith("/070")):
                        ese_index = i
                    elif i == 2 and token.endswith("/100"):
                        total_index = i
                
                if ise_index >= 0 and ise_index < len(processed_marks):
                    key_insem = base_subject + " (Insem)"
                    current_student[key_insem] = processed_marks[ise_index]
                
                if ese_index >= 0 and ese_index < len(processed_marks):
                    key_ese = base_subject + " (ESE)"
                    current_student[key_ese] = processed_marks[ese_index]
                
                if total_index >= 0 and total_index < len(processed_marks):
                    key_total = base_subject + " (Total)"
                    current_student[key_total] = processed_marks[total_index]
                
                # Add grades information
                grade_index = -1
                gp_index = -1
                cp_index = -1
                
                for i, token in enumerate(tokens):
                    if token == "Grd" and i+1 < len(tokens):
                        grade_index = i+1
                    elif token == "GP" and i+1 < len(tokens):
                        gp_index = i+1
                    elif token == "CP" and i+1 < len(tokens):
                        cp_index = i+1
                
                if grade_index >= 0 and grade_index < len(tokens):
                    key_grade = base_subject + " (Grade)"
                    current_student[key_grade] = tokens[grade_index]
                
                if gp_index >= 0 and gp_index < len(tokens):
                    key_gp = base_subject + " (GP)"
                    current_student[key_gp] = tokens[gp_index]
                
                if cp_index >= 0 and cp_index < len(tokens):
                    key_cp = base_subject + " (CP)"
                    current_student[key_cp] = tokens[cp_index]
                
                # Add Tot% column
                tot_percent_index = -1
                for i, token in enumerate(tokens):
                    if token == "Tot%":
                        tot_percent_index = i+1
                        break
                
                if tot_percent_index >= 0 and tot_percent_index < len(tokens):
                    key_tot_percent = base_subject + " (Tot%)"
                    current_student[key_tot_percent] = tokens[tot_percent_index]

    if current_student is not None:
        current_student["CGPA"] = current_student["SGPA"]
        
        # Calculate total marks and percentage for the last student
        total_marks = 0
        max_marks = 0
        for key, value in current_student.items():
            if key.endswith(" (Total)") and value not in ["-", "AB", "FF"]:
                try:
                    total_marks += int(value)
                    max_marks += 100  # Assuming each subject is out of 100
                except ValueError:
                    pass
            elif (key.endswith(" (TW)") or key.endswith(" (PR)")) and value not in ["-", "AB", "FF"]:
                try:
                    total_marks += int(value)
                    max_marks += 50  # Assuming TW/PR are out of 50
                except ValueError:
                    pass
        
        if max_marks > 0:
            current_student["Total"] = str(total_marks)
            current_student["%"] = str(round((total_marks / max_marks) * 100, 2))
        
        students.append(current_student)
    
    return students

def create_excel_in_memory(students, selected_subjects_str):
    """
    Create an Excel file in memory (as bytes) using the student data.
    The selected_subjects_str is a comma-separated list of base subject names.
    Only dynamic keys from each student that match these base names will be included.
    """
    common_cols = ["Seat No.", "Name of Student", "Mother's Name", "PRN", "College Code", 
                  "SGPA", "Total Credits", "CGPA", "Total", "%"]
    subject_keys = set()
    for s in students:
        for k in s.keys():
            if k not in common_cols:
                subject_keys.add(k)
    
    user_subjects = [x.strip() for x in selected_subjects_str.split(",") if x.strip()]
    final_subject_cols = []
    
    # First, handle the user-selected subjects in order
    for subj in user_subjects:
        ordered_keys = []
        # Order the columns: Code, Insem, ESE, Total, TW, PR, etc.
        field_order = [" (Code)", " (Insem)", " (ESE)", " (Total)", " (TW)", " (PR)", " (Status)", 
                      " (Tot%)", " (Grade)", " (GP)", " (CP)"]
        
        for field in field_order:
            for key in subject_keys:
                if key.upper().startswith(subj.upper()) and key.endswith(field):
                    ordered_keys.append(key)
        
        for key in ordered_keys:
            if key not in final_subject_cols:
                final_subject_cols.append(key)
    
    # Add any remaining subject fields not covered by user selection
    for key in sorted(subject_keys):
        if key not in final_subject_cols:
            final_subject_cols.append(key)
    
    final_cols = ["Sr.", "Seat No.", "Name of Student", "Mother's Name", "PRN", "College Code"] + final_subject_cols + ["SGPA", "Total Credits", "CGPA", "Total", "%"]
    
    rows = []
    sr = 1
    for s in students:
        row = {"Sr.": sr}
        for col in final_cols[1:]:
            row[col] = s.get(col, "-")
        rows.append(row)
        sr += 1
    
    df = pd.DataFrame(rows, columns=final_cols)
    
    # Format the Excel file with proper column widths
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Student Data")
        worksheet = writer.sheets["Student Data"]
        
        # Auto-adjust column widths
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, min(max_len, 30))
        
        # Add some basic formatting
        header_format = writer.book.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        worksheet.set_row(0, None, header_format)
        
    return output.getvalue()

# ---------- Multi-Page Streamlit App ----------

def main():
    st.set_page_config(page_title="Result Ledger Parser", layout="wide")
    
    # Sidebar for navigation
    page = st.sidebar.selectbox("Navigation", ["Home", "Contact", "Help"])

    if page == "Home":
        st.title("Dynamic Result Ledger Parser from PDF")
        
        st.markdown("""
        **Instructions:**
        
        1. **Upload a PDF file:**  
           The file should contain result ledger data from Savitribai Phule Pune University.
        
        2. **Auto-Detection of Subjects:**  
           The app will auto-detect subject base names from the PDF.
           You can review and modify these names (comma separated) below.
        
        3. **Generate the Ledger:**  
           Confirm the subject names by checking the box, then click **"Generate Excel"** to process the data.
        """)

        uploaded_pdf = st.file_uploader("Choose a PDF file", type=["pdf"])
        
        if uploaded_pdf is not None:
            pdf_bytes = uploaded_pdf.getvalue()
            extracted_text = extract_text_from_pdf(pdf_bytes)
            
            # Auto-detect subject base names
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
                    with st.spinner("Processing PDF and generating Excel..."):
                        students = parse_student_file_from_text(extracted_text)
                        if students:
                            st.success(f"Successfully extracted data for {len(students)} students.")
                            excel_bytes = create_excel_in_memory(students, subject_names_input)
                            st.download_button(
                                label="Download Excel File",
                                data=excel_bytes,
                                file_name="result_ledger_output.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.error("No student data was extracted. Please check if the PDF format is correct.")

    elif page == "Contact":
        st.title("Contact Information")
        st.markdown("""
        **For Support:**
        
        **Name:** Yash Bari  
        **Email:** [yashbari99@gmail.com](mailto:yashbari99@gmail.com)  
        **Phone:** 9834148536
        
        If you have any questions or need assistance with this application, please feel free to reach out.
        """)

    elif page == "Help":
        st.title("Help & Troubleshooting")
        st.markdown("""
        ### Common Issues and Solutions
        
        1. **Missing Data in Excel Output:**
           - Make sure your PDF is properly scanned and readable
           - Check if the text extraction correctly identifies all records
           - Verify that the subject names are correctly identified
        
        2. **Format Problems:**
           - The parser expects a specific format from SPPU result ledgers
           - If your PDF has a different format, some data might be missed
        
        3. **Subject Detection Issues:**
           - If subjects are not detected automatically, manually enter them
           - Use the exact names as they appear in the PDF (e.g., "DESIGN & ANALYSIS OF ALGO.")
        
        ### PDF Format Requirements
        
        The parser expects PDFs with the following structure:
        
        - Header lines with student details (SEAT NO., NAME, MOTHER, PRN, CLG.)
        - Course lines with subject marks (e.g., "410241 DESIGN & ANALYSIS OF ALGO. * 017/030 AB/070 017/100")
        - SGPA and credits information
        
        ### Need More Help?
        
        If you continue to experience issues, please contact the developer using the information on the Contact page.
        """)

if __name__ == "__main__":
    main()
