import re
import pandas as pd

def clean_mark(token):
    """
    If token contains '/', return only the part before the slash.
    Otherwise, return the token as is.
    """
    if "/" in token:
        return token.split("/")[0]
    return token

def parse_student_file(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    students = []
    current_student = None

    # Create a new student record with default values.
    def new_student():
        return {
            "Seat No.": "-",
            "Name of Student": "-",
            "DAA (Insem)": "-", "DAA (ESE)": "-", "DAA (Total)": "-",
            "ML (Insem)": "-", "ML (ESE)": "-", "ML (Total)": "-",
            "BCT (Insem)": "-", "BCT (ESE)": "-", "BCT (Total)": "-",
            "OOMD (Insem)": "-", "OOMD (ESE)": "-", "OOMD (Total)": "-",
            "STQA (Insem)": "-", "STQA (ESE)": "-", "STQA (Total)": "-",
            "LP-III (TW)": "-", "LP-III (PR)": "-",
            "LP-IV (TW)": "-",
            "PROJECT (TW)": "-",
            "MOOC": "-",
            "SGPA": "-",
            "Total Credits": "-",
            "Total": "",  # Blank, not computed
            "%": "",      # Blank, not computed
            "CGPA": "-"   # Will be set equal to SGPA
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

        # New student header found.
        header_match = header_regex.search(line)
        if header_match:
            if current_student is not None:
                # Set CGPA equal to SGPA.
                current_student["CGPA"] = current_student["SGPA"]
                students.append(current_student)
            current_student = new_student()
            current_student["Seat No."] = header_match.group(1)
            current_student["Name of Student"] = header_match.group(2).strip()
            continue

        # Capture SGPA and Total Credits
        if line.startswith("SGPA1 :"):
            sgpa_match = re.search(r"SGPA1\s*:\s*([\d.]+)", line)
            tc_match = re.search(r"TOTAL CREDITS EARNED\s*:\s*(\d+)", line)
            if sgpa_match:
                current_student["SGPA"] = sgpa_match.group(1)
            if tc_match:
                current_student["Total Credits"] = tc_match.group(1)
            continue

        # Process course lines: they start with a course code (a digit) and contain a "*" token.
        if re.match(r"^\d", line):
            tokens = line.split()
            try:
                star_index = tokens.index("*")
            except ValueError:
                continue  # Skip if the expected structure isn't present

            course_code = tokens[0]
            # Course name: tokens from index 1 up to the "*" token.
            course_name = " ".join(tokens[1:star_index])
            marks = tokens[star_index+1:]
            # Clean all marks tokens: if they are in form X/Y, keep only X.
            marks = [clean_mark(tok) for tok in marks]

            # For theory subjects, we expect at least three marks tokens.
            if course_code == "410241":
                # DESIGN & ANALYSIS OF ALGO.
                if len(marks) >= 3:
                    current_student["DAA (Insem)"] = marks[0]
                    current_student["DAA (ESE)"] = marks[1]
                    current_student["DAA (Total)"] = marks[2]
            elif course_code == "410242":
                # MACHINE LEARNING
                if len(marks) >= 3:
                    current_student["ML (Insem)"] = marks[0]
                    current_student["ML (ESE)"] = marks[1]
                    current_student["ML (Total)"] = marks[2]
            elif course_code == "410243":
                # BLOCKCHAIN TECHNOLOGY
                if len(marks) >= 3:
                    current_student["BCT (Insem)"] = marks[0]
                    current_student["BCT (ESE)"] = marks[1]
                    current_student["BCT (Total)"] = marks[2]
            elif course_code in ["410244", "410244D"]:
                # OBJ. ORIENTED MODL. & DESG.
                if len(marks) >= 3:
                    current_student["OOMD (Insem)"] = marks[0]
                    current_student["OOMD (ESE)"] = marks[1]
                    current_student["OOMD (Total)"] = marks[2]
            elif course_code in ["410245", "410245D"]:
                # SOFT. TEST. & QLTY ASSURANCE
                if len(marks) >= 3:
                    current_student["STQA (Insem)"] = marks[0]
                    current_student["STQA (ESE)"] = marks[1]
                    current_student["STQA (Total)"] = marks[2]
            elif course_code == "410246":
                # LABORATORY PRACTICE - III: expect tokens for TW and PR after "*"
                if len(marks) >= 5:
                    current_student["LP-III (TW)"] = marks[3]
                    current_student["LP-III (PR)"] = marks[4]
            elif course_code == "410247":
                # LABORATORY PRACTICE - IV
                if len(marks) >= 4:
                    current_student["LP-IV (TW)"] = marks[3]
            elif course_code == "410248":
                # PROJECT STAGE - I
                if len(marks) >= 4:
                    current_student["PROJECT (TW)"] = marks[3]
            elif course_code.startswith("410249"):
                # MOOC - LEARN NEW SKILLS
                if len(marks) >= 4:
                    current_student["MOOC"] = marks[3]
            # Skip any duplicate or unrecognized course lines.
    
    if current_student is not None:
        current_student["CGPA"] = current_student["SGPA"]
        students.append(current_student)
    return students

def create_excel(students, output_path):
    cols = [
        "Sr.", "Seat No.", "Name of Student",
        "DAA (Insem)", "DAA (ESE)", "DAA (Total)",
        "ML (Insem)", "ML (ESE)", "ML (Total)",
        "BCT (Insem)", "BCT (ESE)", "BCT (Total)",
        "OOMD (Insem)", "OOMD (ESE)", "OOMD (Total)",
        "STQA (Insem)", "STQA (ESE)", "STQA (Total)",
        "LP-III (TW)", "LP-III (PR)",
        "LP-IV (TW)",
        "PROJECT (TW)",
        "MOOC",
        "Total", "%", "SGPA", "CGPA"
    ]
    
    rows = []
    sr = 1
    for s in students:
        row = {
            "Sr.": sr,
            "Seat No.": s.get("Seat No.", "-"),
            "Name of Student": s.get("Name of Student", "-"),
            "DAA (Insem)": s.get("DAA (Insem)", "-"),
            "DAA (ESE)": s.get("DAA (ESE)", "-"),
            "DAA (Total)": s.get("DAA (Total)", "-"),
            "ML (Insem)": s.get("ML (Insem)", "-"),
            "ML (ESE)": s.get("ML (ESE)", "-"),
            "ML (Total)": s.get("ML (Total)", "-"),
            "BCT (Insem)": s.get("BCT (Insem)", "-"),
            "BCT (ESE)": s.get("BCT (ESE)", "-"),
            "BCT (Total)": s.get("BCT (Total)", "-"),
            "OOMD (Insem)": s.get("OOMD (Insem)", "-"),
            "OOMD (ESE)": s.get("OOMD (ESE)", "-"),
            "OOMD (Total)": s.get("OOMD (Total)", "-"),
            "STQA (Insem)": s.get("STQA (Insem)", "-"),
            "STQA (ESE)": s.get("STQA (ESE)", "-"),
            "STQA (Total)": s.get("STQA (Total)", "-"),
            "LP-III (TW)": s.get("LP-III (TW)", "-"),
            "LP-III (PR)": s.get("LP-III (PR)", "-"),
            "LP-IV (TW)": s.get("LP-IV (TW)", "-"),
            "PROJECT (TW)": s.get("PROJECT (TW)", "-"),
            "MOOC": s.get("MOOC", "-"),
            "Total": s.get("Total", ""),
            "%": s.get("%", ""),
            "SGPA": s.get("SGPA", "-"),
            "CGPA": s.get("CGPA", "-")
        }
        rows.append(row)
        sr += 1

    df = pd.DataFrame(rows, columns=cols)
    df.to_excel(output_path, index=False)
    print(f"Data successfully written to {output_path}")

if __name__ == "__main__":
    input_file = "dat.txt"      # Replace with your file name/path
    output_file = "output.xlsx" # Desired output Excel file name
    students = parse_student_file(input_file)
    create_excel(students, output_file)
