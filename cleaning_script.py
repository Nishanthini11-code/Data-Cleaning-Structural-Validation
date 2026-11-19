
import pandas as pd
import numpy as np

# Step 1: Load the Excel dataset
file_path = "dataset/raw_retail_sales_dataset.xlsx"

df = pd.read_excel(file_path, sheet_name="Raw_Sales_Data")

# Step 2: Display first 5 records
print("\n----- FIRST 5 RECORDS -----")
print(df.head())

# Step 3: Display dataset information
print("\n----- DATASET INFORMATION -----")
print(df.info())

# Step 4: Display number of rows and columns
print("\n----- DATASET SHAPE -----")
print("Rows and Columns:", df.shape)

# Step 5: Check missing values
print("\n----- MISSING VALUES -----")
print(df.isnull().sum())

# Step 6: Check duplicate records
print("\n----- DUPLICATE RECORDS -----")
print("Duplicate rows:", df.duplicated().sum())

# Step 7: Statistical summary
print("\n----- STATISTICAL SUMMARY -----")
print(df.describe(include="all"))

# STEP 7: REMOVE DUPLICATE RECORDS

print("\n----- DUPLICATE REMOVAL -----")

# Count duplicates before removal
duplicates_before = df.duplicated().sum()

print("Duplicates before cleaning:", duplicates_before)

# Remove duplicate records
df = df.drop_duplicates()

# Check records after removal
duplicates_after = df.duplicated().sum()

print("Duplicates after cleaning:", duplicates_after)

print("Records after duplicate removal:", len(df))

# Save the result
df.to_excel("output/deduplicated_sales_data.xlsx", index=False)

print("Deduplicated dataset saved successfully!")

# STEP 8: HANDLE MISSING VALUES

print("\n----- MISSING VALUE ANALYSIS -----")

# Count missing values in each column
print("Missing values before cleaning:")
print(df.isnull().sum())

# Calculate total missing values
total_missing = df.isnull().sum().sum()

print("Total missing values:", total_missing)

# Fill missing categorical values with Unknown

categorical_columns = [
    "Customer_Name",
    "Product",
    "Category",
    "City",
    "Payment_Method"
]

for column in categorical_columns:
    df[column] = df[column].fillna("Unknown")

print("\nMissing text values handled!")
print(df[categorical_columns].isnull().sum())
df.to_excel("output/missing_values_handled.xlsx", index=False)

print("Missing-value dataset saved successfully!")

# STEP 9: STANDARDIZE INCONSISTENT DATA

print("\n----- DATA STANDARDIZATION -----")

# Remove extra spaces and standardize customer names
df["Customer_Name"] = (
    df["Customer_Name"]
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
    .str.title()
)

# Standardize city names
df["City"] = df["City"].str.strip().str.title()

# Standardize product names
df["Product"] = df["Product"].str.strip().str.lower()

product_mapping = {
    "wireless mouse": "Wireless Mouse",
    "bluetooth speaker": "Bluetooth Speaker",
    "notebook": "Notebook",
    "ball pen pack": "Ball Pen Pack",
    "water bottle": "Water Bottle",
    "lunch box": "Lunch Box",
    "cotton t-shirt": "Cotton T-Shirt",
    "backpack": "Backpack",
    "desk lamp": "Desk Lamp",
    "usb cable": "USB Cable"
}

df["Product"] = df["Product"].replace(product_mapping)

# Standardize category names
df["Category"] = df["Category"].str.strip().str.title()

# Standardize payment methods
df["Payment_Method"] = (
    df["Payment_Method"].str.strip().str.lower()
)

payment_mapping = {
    "upi": "UPI",
    "cash": "Cash",
    "credit card": "Credit Card",
    "debit card": "Debit Card",
    "net banking": "Net Banking",
    "unknown": "Unknown"
}

df["Payment_Method"] = df["Payment_Method"].replace(payment_mapping)

print("Data standardization completed!")

print("\nUnique Cities:")
print(df["City"].unique())

print("\nUnique Payment Methods:")
print(df["Payment_Method"].unique())

# Save standardized data
df.to_excel("output/standardized_sales_data.xlsx", index=False)

print("\nStandardized dataset saved successfully!")

# STEP 10: DATE FORMAT STANDARDIZATION

print("\n----- DATE STANDARDIZATION -----")

# Convert dates into a consistent datetime format
df["Order_Date"] = pd.to_datetime(
    df["Order_Date"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

# Count invalid or missing dates
invalid_dates = df["Order_Date"].isnull().sum()

print("Invalid or missing dates:", invalid_dates)

# Display standardized dates
print("\nStandardized Dates:")
print(df["Order_Date"].head(10))

# Save the dataset
df.to_excel("output/date_standardized_data.xlsx", index=False)

print("\nDate standardization completed!")

# STEP 11: NUMERICAL DATA CLEANING

print("\n----- NUMERICAL DATA CLEANING -----")

# Convert Quantity into numeric format
df["Quantity"] = pd.to_numeric(
    df["Quantity"].astype(str).str.strip(),
    errors="coerce"
)

# Convert Unit_Price into numeric format
df["Unit_Price"] = pd.to_numeric(
    df["Unit_Price"]
    .astype(str)
    .str.replace("₹", "", regex=False)
    .str.replace(",", "", regex=False)
    .str.strip(),
    errors="coerce"
)

# Identify invalid values
df.loc[df["Quantity"] <= 0, "Quantity"] = np.nan
df.loc[df["Unit_Price"] <= 0, "Unit_Price"] = np.nan

print("Invalid numerical values identified.")

# Fill missing Quantity with median
quantity_median = df["Quantity"].median()

df["Quantity"] = df["Quantity"].fillna(
    quantity_median
).round().astype(int)

# Fill missing prices using product-wise median
product_median = df.groupby("Product")["Unit_Price"].transform("median")

df["Unit_Price"] = df["Unit_Price"].fillna(product_median)

# Use overall median only if product median is unavailable
df["Unit_Price"] = df["Unit_Price"].fillna(
    df["Unit_Price"].median()
)

print("\nNumerical cleaning completed!")

print("\nMissing numerical values:")
print(df[["Quantity", "Unit_Price"]].isnull().sum())

print("\nNumerical data summary:")
print(df[["Quantity", "Unit_Price"]].describe())

# Save cleaned numerical data
df.to_excel("output/numerical_cleaned_data.xlsx", index=False)

print("\nNumerical dataset saved successfully!")

# STEP 12: STRUCTURAL VALIDATION

print("\n----- STRUCTURAL VALIDATION -----")

# Check required columns
required_columns = [
    "Order_ID", "Customer_Name", "Product",
    "Category", "Quantity", "Unit_Price",
    "Order_Date", "City", "Payment_Method"
]

columns_valid = all(col in df.columns for col in required_columns)

# Check duplicate records
duplicates_valid = df.duplicated().sum() == 0

# Check unique Order IDs
order_id_valid = df["Order_ID"].is_unique

# Check Quantity
quantity_valid = (
    df["Quantity"].notna().all()
    and (df["Quantity"] > 0).all()
    and (df["Quantity"] % 1 == 0).all()
)

# Check Unit Price
price_valid = (
    df["Unit_Price"].notna().all()
    and (df["Unit_Price"] > 0).all()
)

# Check date data type
date_type_valid = pd.api.types.is_datetime64_any_dtype(
    df["Order_Date"]
)

# Check categorical missing values
categorical_valid = df[
    ["Customer_Name", "Product", "Category",
     "City", "Payment_Method"]
].notna().all().all()

# Count dates requiring review
dates_requiring_review = df["Order_Date"].isna().sum()

# Create validation results
validation_results = [
    ["Required columns", columns_valid],
    ["Duplicate removal", duplicates_valid],
    ["Unique Order IDs", order_id_valid],
    ["Valid Quantity", quantity_valid],
    ["Valid Unit Price", price_valid],
    ["Date data type", date_type_valid],
    ["Categorical values", categorical_valid]
]

# Display validation results
for check, result in validation_results:
    status = "PASS" if result else "FAIL"
    print(check, ":", status)

print("Dates requiring review:", dates_requiring_review)

# Create validation report
report = pd.DataFrame(
    validation_results,
    columns=["Validation Check", "Result"]
)

# Save final cleaned dataset
df.to_excel(
    "output/final_cleaned_dataset.xlsx",
    index=False
)

# Save validation report
report.to_excel(
    "output/structural_validation_report.xlsx",
    index=False
)

print("\nFinal cleaned dataset saved!")
print("Structural validation report saved!")
print("Final number of records:", len(df))
print("Final number of columns:", len(df.columns))

# STEP 12: STRUCTURAL VALIDATION

print("\n----- STRUCTURAL VALIDATION -----")

# Check required columns
required_columns = [
    "Order_ID", "Customer_Name", "Product",
    "Category", "Quantity", "Unit_Price",
    "Order_Date", "City", "Payment_Method"
]

columns_valid = all(col in df.columns for col in required_columns)

# Check duplicate records
duplicates_valid = df.duplicated().sum() == 0

# Check unique Order IDs
order_id_valid = df["Order_ID"].is_unique

# Check Quantity
quantity_valid = (
    df["Quantity"].notna().all()
    and (df["Quantity"] > 0).all()
    and (df["Quantity"] % 1 == 0).all()
)

# Check Unit Price
price_valid = (
    df["Unit_Price"].notna().all()
    and (df["Unit_Price"] > 0).all()
)

# Check date data type
date_type_valid = pd.api.types.is_datetime64_any_dtype(
    df["Order_Date"]
)

# Check categorical missing values
categorical_valid = df[
    ["Customer_Name", "Product", "Category",
     "City", "Payment_Method"]
].notna().all().all()

# Count dates requiring review
dates_requiring_review = df["Order_Date"].isna().sum()

# Create validation results
validation_results = [
    ["Required columns", columns_valid],
    ["Duplicate removal", duplicates_valid],
    ["Unique Order IDs", order_id_valid],
    ["Valid Quantity", quantity_valid],
    ["Valid Unit Price", price_valid],
    ["Date data type", date_type_valid],
    ["Categorical values", categorical_valid]
]

# Display validation results
for check, result in validation_results:
    status = "PASS" if result else "FAIL"
    print(check, ":", status)

print("Dates requiring review:", dates_requiring_review)

# Create validation report
report = pd.DataFrame(
    validation_results,
    columns=["Validation Check", "Result"]
)

# Save final cleaned dataset
df.to_excel(
    "output/final_cleaned_dataset.xlsx",
    index=False
)

# Save validation report
report.to_excel(
    "output/structural_validation_report.xlsx",
    index=False
)

print("\nFinal cleaned dataset saved!")
print("Structural validation report saved!")
print("Final number of records:", len(df))
print("Final number of columns:", len(df.columns))

# STEP 13: GENERATE DATA QUALITY REPORT

print("\n----- DATA QUALITY REPORT -----")

# Load original dataset for comparison
raw_df = pd.read_excel(
    "dataset/raw_retail_sales_dataset.xlsx",
    sheet_name="Raw_Sales_Data"
)

# Calculate before-cleaning statistics
raw_rows = len(raw_df)
raw_duplicates = raw_df.duplicated().sum()
raw_missing = raw_df.isnull().sum().sum()

# Count invalid or missing dates in original dataset
raw_dates = pd.to_datetime(
    raw_df["Order_Date"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

raw_invalid_dates = raw_dates.isnull().sum()

# Calculate after-cleaning statistics
final_rows = len(df)
final_duplicates = df.duplicated().sum()
final_missing = df.isnull().sum().sum()
final_invalid_dates = df["Order_Date"].isnull().sum()

# Create overall quality summary
quality_summary = pd.DataFrame({
    "Metric": [
        "Total Records",
        "Duplicate Records",
        "Missing Cells",
        "Invalid or Missing Dates"
    ],
    "Before Cleaning": [
        raw_rows,
        raw_duplicates,
        raw_missing,
        raw_invalid_dates
    ],
    "After Cleaning": [
        final_rows,
        final_duplicates,
        final_missing,
        final_invalid_dates
    ]
})

print("\nBefore and After Cleaning:")
print(quality_summary)

# Create column-wise missing value report
missing_report = pd.DataFrame({
    "Column": raw_df.columns,
    "Missing Before Cleaning": raw_df.isnull().sum().values,
    "Missing After Cleaning": df.isnull().sum().values
})

# Save quality report with multiple sheets
with pd.ExcelWriter(
    "output/data_quality_report.xlsx",
    engine="openpyxl"
) as writer:

    quality_summary.to_excel(
        writer,
        sheet_name="Quality Summary",
        index=False
    )

    missing_report.to_excel(
        writer,
        sheet_name="Missing Values",
        index=False
    )

    report.to_excel(
        writer,
        sheet_name="Validation Results",
        index=False
    )

print("\nData quality report generated successfully!")

# STEP 14: CREATE DATA QUALITY VISUALIZATION

import matplotlib.pyplot as plt

categories = [
    "Duplicate Records",
    "Missing Cells",
    "Invalid Dates"
]

before = [
    raw_duplicates,
    raw_missing,
    raw_invalid_dates
]

after = [
    final_duplicates,
    final_missing,
    final_invalid_dates
]

x = np.arange(len(categories))
width = 0.35

plt.figure(figsize=(9, 5))

plt.bar(
    x - width/2,
    before,
    width,
    label="Before Cleaning"
)

plt.bar(
    x + width/2,
    after,
    width,
    label="After Cleaning"
)

plt.xlabel("Data Quality Issues")
plt.ylabel("Number of Issues")
plt.title("Data Cleaning: Before vs After")
plt.xticks(x, categories, rotation=10)
plt.legend()
plt.tight_layout()

plt.savefig(
    "output/cleaning_summary.png",
    dpi=300
)

plt.show()

print("\nData quality visualization saved successfully!")

