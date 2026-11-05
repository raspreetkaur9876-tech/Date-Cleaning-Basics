# Date Cleaning Basics 📅

## Project Overview
Date Cleaning Basics is a beginner-friendly data cleaning project using Microsoft Excel and Python Pandas. It focuses on standardizing inconsistent date formats and identifying invalid date values.

## Objectives
- Identify inconsistent date formats.
- Convert dates into a standard format.
- Detect invalid dates.
- Maintain an issue log.
- Generate a cleaned Excel output file.

## Tools & Technologies
- Microsoft Excel
- Python 3
- Pandas
- Openpyxl
- Visual Studio Code
- GitHub

## Project Files
| File Name | Description |
|---|---|
| `Retail_Sales_Dates.xlsx` | Original retail sales dataset |
| `clean_dates.py` | Python date cleaning script |
| `Date_Cleaning_Output.xlsx` | Cleaned dataset and issue log |

## Data Cleaning Process
1. Imported the Excel dataset.
2. Checked different date formats.
3. Standardized dates into DD-MM-YYYY format.
4. Identified invalid date values.
5. Created an issue log.
6. Exported the cleaned data into Excel.

## Output
- 8 valid dates
- 2 invalid dates
- Standardized date format: DD-MM-YYYY
- Separate issue log for invalid records.

## How to Run the Project

**Step 1:** Install required libraries.

```bash
python -m pip install pandas openpyxl
```

**Step 2:** Run the Python script.

```bash
python clean_dates.py
```

**Step 3:** Open `Date_Cleaning_Output.xlsx` to view the results.

## Skills Learned
- Data Cleaning
- Date Formatting
- Data Validation
- Python Pandas
- Excel Data Management
- Handling Missing and Invalid Values

## Author
**Raspreet Kaur**

BTech – Artificial Intelligence and Machine Learning

## Conclusion
This project demonstrates how Excel and Python Pandas can be used to clean inconsistent date values, identify invalid records, and prepare structured data for further analysis.
