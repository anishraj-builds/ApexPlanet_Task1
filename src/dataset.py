# ==========================================================
# ApexPlanet Internship - Task 1
# Data Immersion & Wrangling
# Author: Anish Raj.R
# ==========================================================

# -------------------------
# Import Libraries
# -------------------------

import os
import warnings
import logging
import matplotlib
matplotlib.use("Agg")   
import matplotlib.pyplot as plt

logging.basicConfig(
    filename="reports/log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

import pandas as pd

from ydata_profiling import ProfileReport

from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Table,
    TableStyle,
)

warnings.filterwarnings("ignore")

# -------------------------
# Defining Constants
# -------------------------

INPUT_FILE = "data/ApexPlanet_DataAnalytics_Dataset.xlsx"
OUTPUT_FILE = "data/cleaned_sales_dataset.csv"
REPORT_FOLDER = "reports"
CHART_FOLDER = "reports/charts"

LINE = "=" * 60

# -------------------------
# Create Project Folders
# -------------------------

os.makedirs("data", exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)
os.makedirs(CHART_FOLDER, exist_ok=True)

# -------------------------
# Helper Functions
# -------------------------

def log_and_print(message: str) -> None:
    print(message)
    logging.info(message)

# -------------------------
# FUNCTION-1: Load Dataset
# -------------------------
def load_data() -> pd.DataFrame | None:
    """
    Load the Excel dataset and display an overview.

    Returns:
        pd.DataFrame | None:
            The loaded dataset if successful.
            Returns None if the dataset cannot be loaded.
    """

    # -------------------------
    # Load Dataset
    # -------------------------

    print(LINE)
    log_and_print("APEXPLANET DATA ANALYTICS TASK 1")
    print(LINE)

    print("\nLoading Dataset...")

    # -------------------------
    # Read Dataset
    # -------------------------

    try:
        df = pd.read_excel(INPUT_FILE)
        print("Dataset Loaded Successfully.")

    except FileNotFoundError:
        print(f"Error: '{INPUT_FILE}' not found.")
        return None

    except Exception as e:
        print(f"Error loading dataset: {e}")
        return None

    # -------------------------
    # Dataset Overview
    # -------------------------

    print("\nShape of Dataset")
    print(df.shape)

    print("\nRows :", df.shape[0])
    print("Columns :", df.shape[1])

    print("\nColumn Names")

    for col in df.columns:
        print("-", col)

    print("\nData Types")
    print(df.dtypes)

    print("\nFirst Five Rows")
    print(df.head())

    # -------------------------
    # Missing Values
    # -------------------------

    print("\nMissing Values")

    print(df.isnull().sum())

    # -------------------------
    # Duplicate Rows
    # -------------------------

    print("\nDuplicate Rows :", df.duplicated().sum())

    # -------------------------
    # Basic Statistics
    # -------------------------

    print("\nStatistical Summary")

    print(df.describe(include="all"))

    # -------------------------
    # Memory Usage
    # -------------------------

    print("\nMemory Usage")

    print(df.memory_usage(deep=True))

    print("\nPart 1 Completed Successfully.")

    return df

# -------------------------
# Function-2: Clean Data
# ------------------------- 

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean the dataset by handling missing values,
    duplicate records, and validating the data.

    Args:
        df:
            Input dataset.

    Returns:
        pd.DataFrame:
            Cleaned dataset.
    """

    if df is None:
        log_and_print("No dataset available for cleaning.")
        return None
    
    print(LINE)
    log_and_print("PART 2 : DATA CLEANING")
    print(LINE)

    # -------------------------
    # Convert Order_Date
    # -------------------------

    print("\nConverting Order_Date to datetime...")

    df["Order_Date"] = pd.to_datetime(
        df["Order_Date"],
        format="%Y-%m-%d",
        errors="coerce"
    )

    log_and_print("Order_Date converted successfully.")

    # -------------------------
    # Handle Missing Values
    # -------------------------

    log_and_print("\nHandling Missing Values...")

    age_missing_before = df["Age"].isna().sum()
    city_missing_before = df["City"].isna().sum()

    median_age = df["Age"].median()
    mode_city = df["City"].mode()[0]

    df["Age"] = df["Age"].fillna(median_age)
    df["City"] = df["City"].fillna(mode_city)

    print(f"Age missing values before : {age_missing_before}")
    print(f"Age missing values after  : {df['Age'].isna().sum()}")

    print(f"\nCity missing values before : {city_missing_before}")
    print(f"City missing values after  : {df['City'].isna().sum()}")

    logging.info("Missing values handled successfully.")

    # -------------------------
    # Remove Duplicate Rows
    # -------------------------

    log_and_print("\nRemoving Duplicate Rows...")

    duplicates_before = df.duplicated().sum()

    df = df.drop_duplicates().reset_index(drop=True)

    duplicates_after = df.duplicated().sum()

    print(f"Duplicate rows before : {duplicates_before}")
    print(f"Duplicate rows after  : {duplicates_after}")

    logging.info(
        f"Duplicate rows before: {duplicates_before}, "
        f"after: {duplicates_after}"
    )

    # -------------------------
    # Remove Duplicate Order_ID
    # -------------------------

    log_and_print("\nChecking Duplicate Order_ID...")

    duplicate_order_ids = df[
        df.duplicated(subset="Order_ID", keep=False)
    ]

    duplicate_count = duplicate_order_ids["Order_ID"].nunique()

    print(f"Duplicate Order_IDs found : {duplicate_count}")

    if duplicate_count > 0:

        print("\nDuplicate Order IDs:")

        print(
            duplicate_order_ids[["Order_ID"]]
            .drop_duplicates()
            .to_string(index=False)
        )

        rows_before = len(df)

        df = df.drop_duplicates(
            subset="Order_ID",
            keep="first"
        ).reset_index(drop=True)

        rows_after = len(df)

        print("\nDuplicate Order_ID records removed.")

        print(f"Rows before cleaning : {rows_before}")
        print(f"Rows after cleaning  : {rows_after}")
        print(f"Rows removed         : {rows_before - rows_after}")
        
        logging.info(
            f"Duplicate Order_ID cleanup completed. "
            f"Rows before: {rows_before}, "
            f"Rows after: {rows_after}, "
            f"Rows removed: {rows_before - rows_after}"
        )

    else:

        log_and_print("No duplicate Order_ID found.")

    # -------------------------
    # Clean Text Columns
    # -------------------------

    log_and_print("\nCleaning Text Columns...")

    text_columns = [
        "Customer_Name",
        "Gender",
        "City",
        "Product",
        "Category"
    ]

    for column in text_columns:
        df[column] = df[column].astype(str).str.strip()

    log_and_print("Extra spaces removed from text columns.")

    # -------------------------
    # Standardize Text
    # -------------------------

    for column in text_columns:
        df[column] = df[column].str.title()

    log_and_print("Text standardized successfully.")

    # -------------------------
    # Validate Total_Sales
    # -------------------------

    log_and_print("\nValidating Total_Sales...")

    df["Calculated_Total"] = (
        df["Quantity"] * df["Unit_Price"]
    ).round(2)

    invalid_sales = df[
        abs(df["Calculated_Total"] - df["Total_Sales"]) > 0.01
    ]

    log_and_print(
        f"Rows with incorrect Total_Sales : {len(invalid_sales)}"
    )

    logging.info(
        "Total_Sales validation completed. Invalid rows: %d",
        len(invalid_sales)
    )

    # -------------------------
    # Remove Helper Column
    # -------------------------

    df.drop(
        columns=["Calculated_Total"],
        inplace=True
    )

    log_and_print("Calculated_Total helper column removed.")

    # -------------------------
    # Final Data Types
    # -------------------------

    print("\nUpdated Data Types")
    print(df.dtypes)

    logging.info(
        "Final data types:\n%s",
        df.dtypes.to_string()
    )

    logging.info(
        f"Final dataset contains {len(df)} rows and "
        f"{len(df.columns)} columns."
    )

    # -------------------------
    # Save Clean Dataset
    # -------------------------

    try:
        df.to_csv(
            OUTPUT_FILE,
            index=False
        )

    except Exception:

        logging.exception(
            "Failed to save cleaned dataset."
        )

        log_and_print(
            "Failed to save cleaned dataset. "
            "Check reports/log.txt for details."
        )

    else:

        log_and_print("Clean dataset saved successfully.")
        log_and_print(f"Location : {OUTPUT_FILE}")

    log_and_print("Part 2 Completed Successfully.")

    return df

# -------------------------
# Function-3: Feature Engineering
# -------------------------

def feature_engineering(df: pd.DataFrame) -> pd.DataFrame | None:
    """
    Create new features from the cleaned dataset.

    Args:
        df:
            Cleaned dataset.

    Returns:
        pd.DataFrame | None:
            Dataset with engineered features.
            Returns None if no dataset is available.
    """

    if df is None:
        log_and_print("No dataset available for feature engineering.")
        return None

    print(LINE)
    log_and_print("PART 3 : FEATURE ENGINEERING")
    print(LINE)

    # -------------------------
    # Extract Year
    # -------------------------

    log_and_print("\nCreating Year column...")

    df["Year"] = df["Order_Date"].dt.year

    log_and_print("Year column created successfully.")

    # -------------------------
    # Extract Month
    # -------------------------

    log_and_print("\nCreating Month column...")

    df["Month"] = df["Order_Date"].dt.month_name()

    log_and_print("Month column created successfully.")

    # -------------------------
    # Extract Quarter
    # -------------------------

    log_and_print("\nCreating Quarter column...")

    df["Quarter"] = "Q" + df["Order_Date"].dt.quarter.astype(str)

    log_and_print("Quarter column created successfully.")

    # -------------------------
    # Extract Day Name
    # -------------------------

    log_and_print("\nCreating Day_Name column...")

    df["Day_Name"] = df["Order_Date"].dt.day_name()

    log_and_print("Day_Name column created successfully.")

    # -------------------------
    # Create Age Group
    # -------------------------

    log_and_print("\nCreating Age_Group column...")

    df["Age_Group"] = pd.cut(
        df["Age"],
        bins=[0, 18, 30, 45, 60, 100],
        labels=[
            "Below 18",
            "19-30",
            "31-45",
            "46-60",
            "60+"
        ]
    )

    log_and_print("Age_Group column created successfully.")

    # -------------------------
    # Create Sales Category
    # -------------------------

    log_and_print("\nCreating Sales_Category column...")

    df["Sales_Category"] = pd.cut(
        df["Total_Sales"],
        bins=[0, 1000, 5000, 10000, float("inf")],
        labels=[
            "Low",
            "Medium",
            "High",
            "Premium"
        ]
    )

    log_and_print("Sales_Category column created successfully.")

    # -------------------------
    # Display New Columns
    # -------------------------

    log_and_print("\nNew Columns Added")

    new_columns = [
        "Year",
        "Month",
        "Quarter",
        "Day_Name",
        "Age_Group",
        "Sales_Category"
    ]

    for column in new_columns:
        log_and_print(f"- {column}")

    logging.info(
        "Feature engineered columns: %s",
        ", ".join(new_columns)
    )

    # -------------------------
    # Save Updated Dataset
    # -------------------------

    try:

        df.to_csv(
            OUTPUT_FILE,
            index=False
        )

    except Exception:

        logging.exception(
            "Failed to save feature engineered dataset."
        )

        log_and_print(
            "Failed to save feature engineered dataset."
        )

    else:

        log_and_print(
            "Feature engineered dataset saved successfully."
        )

        log_and_print(f"Location : {OUTPUT_FILE}")

    log_and_print("Part 3 Completed Successfully.")

    return df

# -------------------------
# Function-4: Generate Reports
# -------------------------

def generate_report(df: pd.DataFrame) -> None:
    """
    Generate a PDF report containing important dataset summaries.
    """

    if df is None:
        log_and_print("No dataset available for report generation.")
        return

    print(LINE)
    log_and_print("PART 4 : REPORT GENERATION")
    print(LINE)

    report_file = os.path.join(
        REPORT_FOLDER,
        "Data_Analysis_Report.pdf"
    )

    document = SimpleDocTemplate(report_file)
    styles = getSampleStyleSheet()
    elements = []

    # -------------------------
    # Report Title
    # -------------------------

    elements.append(
        Paragraph(
            "ApexPlanet Data Analytics Report",
            styles["Title"]
        )
    )

    elements.append(
        Paragraph("<b>ApexPlanet Internship</b>", styles["Heading2"])
    )

    elements.append(
        Paragraph("Task 1 : Data Immersion & Wrangling", styles["Normal"])
    )

    elements.append(
        Paragraph("Author : Anish Raj.R", styles["Normal"])
    )

    elements.append(
        Paragraph(
            f"Generated On : {datetime.now().strftime('%d-%m-%Y')}",
            styles["Normal"]
        )
    )

    elements.append(Paragraph("<br/>", styles["Normal"]))

    # -------------------------
    # Dataset Overview
    # -------------------------

    elements.append(
        Paragraph("<b>Dataset Overview</b>", styles["Heading2"])
    )

    overview_data = [
        ["Metric", "Value"],
        ["Total Rows", str(len(df))],
        ["Total Columns", str(len(df.columns))],
        ["Missing Values", str(df.isna().sum().sum())],
        ["Duplicate Rows", str(df.duplicated().sum())]
    ]

    overview_table = Table(overview_data)

    overview_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 8)
        ])
    )

    elements.append(overview_table)
    elements.append(Paragraph("<br/>", styles["Normal"]))

    # -------------------------
    # Statistical Summary
    # -------------------------

    elements.append(
        Paragraph("<b>Statistical Summary</b>", styles["Heading2"])
    )

    summary = df.describe().round(2)
    summary_data = [["Statistic"] + list(summary.columns)]

    for index, row in summary.iterrows():
        summary_data.append(
            [str(index)] + [str(value) for value in row]
        )

    summary_table = Table(summary_data)

    summary_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTSIZE", (0, 0), (-1, -1), 7)
        ])
    )

    elements.append(summary_table)
    elements.append(Paragraph("<br/>", styles["Normal"]))

    # -------------------------
    # Business Insights
    # -------------------------

    elements.append(
        Paragraph("<b>Business Insights</b>", styles["Heading2"])
    )

    top_category = (
        df.groupby("Category")["Total_Sales"]
        .sum()
        .idxmax()
    )

    top_city = (
        df.groupby("City")["Total_Sales"]
        .sum()
        .idxmax()
    )

    insights = [
        f"• Total Records : {len(df)}",
        f"• Missing Values : {df.isna().sum().sum()}",
        f"• Duplicate Rows : {df.duplicated().sum()}",
        f"• Highest Revenue Category : {top_category}",
        f"• Highest Revenue City : {top_city}",
        "• Dataset cleaned and ready for analysis."
    ]

    for insight in insights:
        elements.append(
            Paragraph(insight, styles["Normal"])
        )

    try:
        document.build(elements)
        logging.info("PDF report created successfully.")

    except Exception:
        logging.exception("Failed to generate PDF report.")
        log_and_print("Failed to generate PDF report.")

    else:
        log_and_print("PDF report generated successfully.")
        log_and_print(f"Location : {report_file}")

    log_and_print("Part 4 Completed Successfully.")

# -------------------------
# Function-5: Create Visualizations
# -------------------------

def create_visualizations(df: pd.DataFrame) -> None:
    """
    Create and save visualizations from the feature
    engineered dataset.

    Args:
        df:
            Feature engineered dataset.

    Returns:
        None
    """

    # -------------------------
    # Validate Dataset
    # -------------------------

    if df is None:
        log_and_print(
            "No dataset available for visualization."
        )
        return

    print(LINE)
    log_and_print("PART 5 : DATA VISUALIZATION")
    print(LINE)

    # =====================================================
    # Chart 1 : Total Sales by Category
    # =====================================================

    log_and_print(
        "\nCreating Sales by Category chart..."
    )

    category_sales = (
        df.groupby("Category")["Total_Sales"]
        .sum()
        .sort_values(ascending=False)
    )

    plt.figure(figsize=(8, 5))

    category_sales.plot(kind="bar")

    plt.title("Total Sales by Category")
    plt.xlabel("Category")
    plt.ylabel("Total Sales")
    plt.xticks(rotation=45)

    chart_path = os.path.join(
        CHART_FOLDER,
        "sales_by_category.png"
    )

    plt.savefig(
        chart_path,
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    log_and_print(
        "Sales by Category chart created successfully."
    )

    # =====================================================
    # Chart 2 : Total Sales by City
    # =====================================================

    log_and_print(
        "\nCreating Sales by City chart..."
    )

    city_sales = (
        df.groupby("City")["Total_Sales"]
        .sum()
        .sort_values(ascending=False)
    )

    plt.figure(figsize=(10, 5))

    city_sales.plot(kind="bar")

    plt.title("Total Sales by City")
    plt.xlabel("City")
    plt.ylabel("Total Sales")
    plt.xticks(rotation=45)

    chart_path = os.path.join(
        CHART_FOLDER,
        "sales_by_city.png"
    )

    plt.savefig(
        chart_path,
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    log_and_print(
        "Sales by City chart created successfully."
    )

    # =====================================================
    # Chart 3 : Monthly Sales Trend
    # =====================================================

    log_and_print(
        "\nCreating Monthly Sales Trend chart..."
    )

    monthly_sales = (
        df.groupby("Month")["Total_Sales"]
        .sum()
    )

    month_order = [
        "January", "February", "March",
        "April", "May", "June",
        "July", "August", "September",
        "October", "November", "December"
    ]

    monthly_sales = (
        monthly_sales
        .reindex(month_order)
        .dropna()
    )

    plt.figure(figsize=(10, 5))

    monthly_sales.plot(kind="line", marker="o")

    plt.title("Monthly Sales Trend")
    plt.xlabel("Month")
    plt.ylabel("Total Sales")
    plt.xticks(rotation=45)
    chart_path = os.path.join(
        CHART_FOLDER,
        "monthly_sales_trend.png"
    )

    plt.savefig(
        chart_path,
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    log_and_print(
        "Monthly Sales Trend chart created successfully."
    )

    # =====================================================
    # Chart 4 : Gender Distribution
    # =====================================================

    log_and_print(
        "\nCreating Gender Distribution chart..."
    )

    gender_count = (
        df["Gender"]
        .value_counts()
    )

    plt.figure(figsize=(6, 6))

    gender_count.plot(
        kind="pie",
        autopct="%1.1f%%"
    )

    plt.ylabel("")
    plt.title("Gender Distribution")

    chart_path = os.path.join(
        CHART_FOLDER,
        "gender_distribution.png"
    )

    plt.savefig(
        chart_path,
        dpi=300,
        bbox_inches="tight"
    )
    plt.close()

    log_and_print(
        "Gender Distribution chart created successfully."
    )

    log_and_print(
        "All visualizations created successfully."
    )

    log_and_print(
        f"Location : {CHART_FOLDER}"
    )

    log_and_print(
        "Part 5 Completed Successfully."
    )

# -------------------------
# Function-6: Generate Profiling Report
# -------------------------

def generate_profiling_report(df):

    if df is None:
        return

    print("\nGenerating Profiling Report...")

    profile = ProfileReport(
        df.drop(columns=[
            "Order_ID",
            "Customer_ID",
            "Customer_Name"
        ]),
        title="Sales Dataset Profiling Report",
        explorative=True
    )

    profile.to_file("reports/Profiling_Report.html")

    print("Profiling Report Generated Successfully.")