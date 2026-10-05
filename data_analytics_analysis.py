"""
IBM SkillsBuild / AICTE Internship
Data Analytics with AI - Prompt Workbook Analysis

Student: Garvita Chauhan
Roll No.: 25MCAN0198
College: JECRC University, Jaipur

This script reads the Raw_Data sheet from the accompanying Excel workbook
and reproduces the core checks used for Prompts 1-3. Prompt 4 contains
management questions for further analysis.

Place this .py file in the same folder as:
    IBM_Data_Analytics_Prompt_Answers_SUBMISSION_READY.xlsx

Run:
    python data_analytics_analysis.py
"""

from pathlib import Path
import pandas as pd

WORKBOOK = Path("IBM_Data_Analytics_Prompt_Answers_SUBMISSION_READY.xlsx")
OUTPUT = Path("analysis_output.xlsx")

REQUIRED_COLUMNS = [
    "Order_ID", "Order_Date", "Customer_ID", "Product",
    "Category", "Region", "Quantity", "Revenue", "Profit"
]

def load_data():
    if not WORKBOOK.exists():
        raise FileNotFoundError(
            f"Workbook not found: {WORKBOOK.resolve()}"
        )
    df = pd.read_excel(WORKBOOK, sheet_name="Raw_Data")
    df = df[REQUIRED_COLUMNS].copy()
    return df

def prompt_1_structure(df):
    rows = []
    for col in df.columns:
        dtype = str(df[col].dtype)
        if col.endswith("_ID"):
            data_type = "Identifier"
        elif col == "Order_Date":
            data_type = "Date"
        elif pd.api.types.is_numeric_dtype(df[col]):
            data_type = "Numerical"
        else:
            data_type = "Categorical"
        rows.append({
            "Column": col,
            "Meaning": {
                "Order_ID": "Unique order identifier",
                "Order_Date": "Date when the order was placed",
                "Customer_ID": "Unique customer identifier",
                "Product": "Product purchased",
                "Category": "Product category",
                "Region": "Region of order",
                "Quantity": "Units ordered",
                "Revenue": "Sales generated",
                "Profit": "Profit generated",
            }[col],
            "Data Type": data_type,
            "Pandas Type": dtype,
        })
    return pd.DataFrame(rows)

def prompt_2_quality(df):
    result = {
        "Missing Customer_ID": int(df["Customer_ID"].isna().sum()),
        "Missing Quantity": int(df["Quantity"].isna().sum()),
        "Missing Region": int(df["Region"].isna().sum()),
        "Negative Quantity": int((pd.to_numeric(df["Quantity"], errors="coerce") < 0).sum()),
        "Negative Profit": int((pd.to_numeric(df["Profit"], errors="coerce") < 0).sum()),
        "Duplicate Order_ID rows": int(df["Order_ID"].duplicated().sum()),
        "Revenue stored as text": int(
            df["Revenue"].map(lambda x: isinstance(x, str)).sum()
        ),
    }
    return pd.DataFrame(
        [{"Check": k, "Result": v} for k, v in result.items()]
    )

def prompt_3_customer_segment(df):
    out = df[["Customer_ID"]].copy()
    counts = df["Customer_ID"].value_counts(dropna=False)
    out["Customer_Segment"] = df["Customer_ID"].map(
        lambda x: "New Customer"
        if pd.isna(x)
        else ("New Customer" if counts.get(x, 0) == 1 else "Repeat Customer")
    )
    return out

def prompt_4_questions():
    questions = [
        ("Sales", "Which month generated the highest total Revenue, and how did monthly Revenue change across the year?"),
        ("Sales", "Which month generated the highest total Profit, and which generated the lowest?"),
        ("Sales", "Which Year/Month combination has the highest order-value mix?"),
        ("Product", "Which Product generated the highest total Revenue?"),
        ("Product", "Which Product generated the highest total Profit and which generated the lowest?"),
        ("Product", "Which Category has the highest total Quantity sold?"),
        ("Customer", "Which customers placed the most orders, and are they classified as Repeat Customers?"),
        ("Customer", "Which Customer_Segment generates more total Revenue: New Customer or Repeat Customer?"),
        ("Customer", "How many high-value orders fall into each Customer_Segment?"),
        ("Region", "Which Region generated the highest total Revenue and which generated the lowest?"),
        ("Region", "Which Region generated the highest total Profit and which had the lowest/negative profitability?"),
        ("Region", "Which Region has the highest average Quantity per order, and how does that compare with its Revenue and Profit?"),
    ]
    return pd.DataFrame(questions, columns=["Category", "Business Question"])

def main():
    df = load_data()

    structure = prompt_1_structure(df)
    quality = prompt_2_quality(df)
    segments = prompt_3_customer_segment(df)
    questions = prompt_4_questions()

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        structure.to_excel(writer, sheet_name="Prompt_1_Structure", index=False)
        quality.to_excel(writer, sheet_name="Prompt_2_Quality", index=False)
        segments.to_excel(writer, sheet_name="Prompt_3_Segments", index=False)
        questions.to_excel(writer, sheet_name="Prompt_4_Questions", index=False)

    print("Analysis completed successfully.")
    print(f"Rows analysed: {len(df)}")
    print(f"Output written to: {OUTPUT.resolve()}")

if __name__ == "__main__":
    main()
