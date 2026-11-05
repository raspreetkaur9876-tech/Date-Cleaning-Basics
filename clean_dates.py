import pandas as pd
from pathlib import Path
from datetime import datetime

# Keep files in the same folder as this Python script
folder = Path(__file__).resolve().parent
input_file = folder / "Retail_Sales_Dates.xlsx"
output_file = folder / "Date_Cleaning_Output.xlsx"

# Read the Excel workbook using the openpyxl engine
df = pd.read_excel(input_file, sheet_name="Cleaned_Data", engine="openpyxl")

# Convert each date carefully, supporting the formats in this project
def clean_date(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value)

    value = str(value).strip()
    formats = [
        "%Y-%m-%d",       # 2026-01-15
        "%d/%m/%Y",       # 15/02/2026
        "%m-%d-%Y",       # 03-18-2026
        "%B %d, %Y",      # March 20, 2026
        "%Y/%m/%d",       # 2026/04/25
    ]
    for date_format in formats:
        try:
            return pd.Timestamp(datetime.strptime(value, date_format))
        except ValueError:
            pass
    return pd.NaT

# Create cleaned dates and validation status
df["Clean_Date"] = df["Order_Date"].apply(clean_date)
df["Issue_Status"] = df["Clean_Date"].apply(
    lambda date: "Valid" if pd.notna(date) else "Check Date"
)

# Create an issue log for dates that could not be converted
issue_log = df.loc[df["Clean_Date"].isna(), ["Order_ID", "Order_Date", "Issue_Status"]].copy()
issue_log.rename(columns={"Order_Date": "Original_Date", "Issue_Status": "Issue"}, inplace=True)

# Display dates consistently in the output workbook
df["Clean_Date"] = df["Clean_Date"].dt.strftime("%d-%m-%Y")

# Save both sheets to a new Excel workbook
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
    issue_log.to_excel(writer, sheet_name="Issue_Log", index=False)

print("Date cleaning completed successfully!")
print("\nCleaned Data:")
print(df.to_string(index=False))
print("\nIssue Log:")
print(issue_log.to_string(index=False))
print(f"\nOutput saved to: {output_file}")
