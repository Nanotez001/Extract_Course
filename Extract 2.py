import fitz  # pip install pymupdf
import re
import pandas as pd
import unicodedata
from collections import Counter

def clean_text(text):
    if not text:
        return ""
    # Normalize text (NFKC)
    text = unicodedata.normalize("NFKC", text)
    # Remove control/private-use characters
    text = "".join(c for c in text if unicodedata.category(c)[0] != "C")
    return text.strip()
def split_detail(detail):
    lines = detail.strip().splitlines()
    detail_th = []
    detail_en = []
    for line in lines:
        if re.search(r'[\u0E00-\u0E7F]', line):
            detail_th.append(line.strip())
        elif line.strip():
            detail_en.append(line.strip())
    return '\n'.join(detail_th), '\n'.join(detail_en)

pua_to_thai = {
    "\uf70a": "่",  # map your PUA char to real Thai character
    "\uf70b": "้",
    "\uf70c": "๊",
    "\uf70d": "๋",
    "\uf70e": "์",
    "\uf703": "ึ",
    "\uf712": "็",
    "\uf705": "่",
    "\uf710": "ั",
    "\uf706": "้",
    "\uf701": "ิ"
    # ... add all you need
}

doc = fitz.open("C:/Users/LEGION by Lenovo/Desktop/TQFSCMA.pdf")

pages_to_process = range(49, 100)  # Your page range (48, 49, 50)

# ==============================================================================
# Step 1: Automatically find common header lines
# ==============================================================================
line_counts = Counter()
# Let's check at least 2 pages to be sure it's a repeating header
num_pages_to_check = max(2, len(list(pages_to_process))) 

# First, we loop through just to count the lines and find headers
for page_num in pages_to_process:
    # PyMuPDF pages are 0-indexed, so we subtract 1
    page = doc.load_page(page_num - 1) 
    lines = page.get_text("text").splitlines()
    
    # We assume headers are within the first 5 lines of a page.
    # You can adjust this number if needed.
    for line in lines[:8]:
        stripped_line = line.strip()
        if stripped_line:  # Ignore empty lines
            line_counts[stripped_line] += 1

# A header is a line that appears on most (or all) pages.
# We create a set of these lines for fast checking later.
header_lines = {line for line, count in line_counts.items() if count >= num_pages_to_check}

# This print statement is useful for debugging to see what headers were found!
print(f"✅ Automatically identified headers: {header_lines}")


# ==============================================================================
# Step 2: Clean the text using the identified headers and extract data
# ==============================================================================
text = ""
for page_num in pages_to_process:
    # print(f"==== Processing page {page_num} ====")
    page = doc.load_page(page_num - 1)
    lines_on_page = page.get_text("text").splitlines()
    
    filtered_lines = []
    for line in lines_on_page:
        stripped_line = line.strip()
        
        # Rule 1: Skip the line if it's one of our automatically found headers
        if stripped_line in header_lines:
            continue
            
        # Rule 2: Skip the line if it looks like a page number (Thai or Arabic)
        if re.fullmatch(r'\s*[๐-๙0-9]+\s*', stripped_line):
            continue
            
        # If the line is not a header or page number, we keep it
        filtered_lines.append(line)
        
    cleaned_page_text = "\n".join(filtered_lines)
    text += cleaned_page_text + "\n"


pattern = re.compile(
    r"(?P<credit>[๐-๙0-9]\([๐-๙0-9\-–]+\))\s*"                                   # credit เช่น ๓(๒-๒-๕)
    r"[\r\n]+(?P<code>[A-Z]{2,4}\s*\d{3})\s*"                                     # code เช่น LAEN 103
    r"[\r\n]+(?P<title>[A-Za-z\u0E00-\u0E7F0-9\s\-\(\)\&\.]+)\s*"                 # title เช่น English Level I
    r"[\r\n]+(?:(?P<prereq_th>วิชาบังคับ[^\r\n]*)[\r\n]+)?"                      # optional Thai prereq
    r"(?:(?P<prereq_en>Prerequisite[^\r\n]*)[\r\n]+)?"                            # optional English prereq
    r"(?P<detail>.+?)(?=(?:[\r\n]+(?:[๐-๙0-9]\([๐-๙0-9\-–]+\)|[ก-ฮ]{4}\s*\d{3})|\Z))",               # detail (จนถึง credit ถัดไปหรือจบไฟล์)
    re.DOTALL
)

# pattern = re.compile(
#     r"(?P<code_th>[ก-ฮ]{4}\s*\d{3})\s*"                                   # e.g. กขคง 101
#     r"[\r\n]+(?P<title_th>[\u0E00-\u0E7F0-9\s\-\(\)\&\.]+)\s*"            # e.g. ภาษาอังกฤษเบื้องต้น
#     r"[\r\n]+(?P<credit>[๐-๙0-9]\([๐-๙0-9\-–]+\))\s*"                    # e.g. ๓(๒-๒-๕)
#     r"[\r\n]+(?P<code_eng>[A-Z]{4}\s*\d{3})\s*"                           # e.g. LAEN 103
#     r"[\r\n]+(?P<title_eng>[A-Za-z0-9\s\-\(\)\&\.]+)"                
#     r"[\r\n]+(?:(?P<prereq_th>วิชาบังคับ[^\r\n]*)[\r\n]+)?"                      # optional Thai prereq
#     r"(?:(?P<prereq_en>Prerequisite[^\r\n]*)[\r\n]+)?"                            # optional English prereq
#     r"(?P<detail>.+?)(?=(?:[\r\n]+[๐-๙0-9]\([๐-๙0-9\-–]+\)|\Z))",                # detail (จนถึง credit ถัดไปหรือจบไฟล์)
#     re.DOTALL
# )


# matches = pattern.findall(text)

# print(f"Total courses found: {len(matches)}")
# for i in range(len(matches)):
#     print(f"Course {i+1}: {matches[i][1]}")
#     # print(matches[i])

for i, m in enumerate(pattern.finditer(text), start=1):
    groups = m.groups()
    clean = [g.strip() for g in groups if g and g.strip()]  # <── ล้างและตัด group ว่างออก
    # print(f"Course {i}:", clean)
    print(f"Course {i}: {m.group('code')}")
    # print(clean)

rows = []
for m in pattern.finditer(text):
    # data = {k: clean_text(v) for k, v in m.groupdict().items()}
    data = {k: (v.strip() if v else "") for k, v in m.groupdict().items()}
    rows.append(data)
print(rows[15])
df = pd.DataFrame(rows)

for pua, thai in pua_to_thai.items():
    df["prereq_th"] = df["prereq_th"].str.replace(pua, thai)
    df["detail"] = df["detail"].str.replace(pua, thai)

# print(df)
# print(df["detail"][0:5])
# print(df.loc[df["code"] == "LAEN 106", "detail"].iloc[0])

df.to_excel(
    "C:/Users/LEGION by Lenovo/Desktop/courses_extracted.xlsx", 
    index=False,
    engine='openpyxl'
)

# Save to CSV with UTF-8 with BOM for Excel
df.to_csv("C:/Users/LEGION by Lenovo/Desktop/courses_extracted.csv", index=False, encoding="utf-8-sig")
