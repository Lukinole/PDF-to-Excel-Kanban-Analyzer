import re
import pandas as pd
from PyPDF2 import PdfReader

# Path to your PDF file
pdf_path = r"C:\Users\User\Desktop\conventor\sourse_file.pdf"

# Load and extract text from all pages
reader = PdfReader(pdf_path)
full_text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])

# Regular expression to extract:
# - Part Number (e.g. 040-K774-A000)
# - Kanban code (e.g. K774)
# - Lot quantity (e.g. 4800 PC)
pattern = r"\(P\)\s+(040-[A-Z0-9]{4}-[A-Z0-9]{4})[\s\S]*?\(K\)\s+([A-Z0-9]{3,})[\s\S]*?Description[\s\S]*?(\d+)\s+PC"

# Find all matching groups
matches = re.findall(pattern, full_text)

# Convert to DataFrame
df = pd.DataFrame(matches, columns=["Part Number", "Kanban", "Lot"])
df["Lot"] = df["Lot"].astype(int)
df["Labels Count"] = 1  # each match represents one label

# Group by Part Number, Kanban, and Lot size (to keep original lot amount)
grouped = df.groupby(["Part Number", "Kanban", "Lot"], as_index=False).agg({"Labels Count": "sum"})

# Calculate total quantity = Lot size * number of labels
grouped["Total Qty"] = grouped["Lot"] * grouped["Labels Count"]

# Save to Excel
output_file = r"C:\Users\User\Desktop\conventor\kanban_summary_final.xlsx"
grouped.to_excel(output_file, index=False)

print("Done! Excel file saved at:", output_file)

