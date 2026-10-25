import re
from pathlib import Path
import pandas as pd


# =========================================================
# 1. FILE PATHS
# =========================================================

file_path = Path(
    r"/Users/ibrahimbarqawi/Desktop/Data_Analysis_Course_and_projects/Codes/17-online_retail_II/Data/online_retail_II.xlsx"
)

output_path = (
    file_path.parent
    /"Cleaned"/"online_retail_II_cleaned.xlsx"
)

if not file_path.exists():
    raise FileNotFoundError(
        f"File not found:\n{file_path}"
    )



config_path = Path(
   "/Users/ibrahimbarqawi/Desktop/Data_Analysis_Course_and_projects/Codes/17-online_retail_II/Model/Cleaning/Verification_Configuration.xlsx"
)

if not config_path.exists():
    raise FileNotFoundError(
        f"Configuration file not found:\n"
        f"{config_path}"
    )



# =========================================================
# 2. READ CONFIGURATION FILE
# =========================================================

configuration_df = pd.read_excel(
    config_path,
    sheet_name="Column Rules"
)

# Remove extra spaces from configuration headers.
configuration_df.columns = (
    configuration_df.columns
    .astype("string")
    .str.strip()
)

required_configuration_columns = ["Sheet", "Column","Data Type"]



missing_configuration_columns = [
    col
    for col in required_configuration_columns
    if col not in configuration_df.columns
]

if missing_configuration_columns:
    raise ValueError(
        "Missing columns in Configuration.xlsx: "
        f"{missing_configuration_columns}"
    )

# Standardize data-type labels.
configuration_df["Data Type"] = (
configuration_df["Data Type"]
.astype("string")
.str.strip()
.str.lower()
)
# Connect Excel data types to the names used by your cleaning code.
configuration_types = {
"string": "string_columns",
"numeric": "numeric_columns",
"date": "date_columns",
"category": "category_columns"
}

# Build the configuration dictionary automatically.
sheet_configurations = {}

for sheet_name, sheet_rows in configuration_df.groupby(
    "Sheet",
    sort=False
):

    sheet_name = str(sheet_name).strip()

    sheet_configurations[sheet_name] = {}

    for data_type , configuration_type in configuration_types.items() :

        configured_columns = ( sheet_rows.loc[sheet_rows["Data Type"] == data_type ,
                                              "Column"
                                              ].dropna().astype("string").str.replace(r"[\x00-\x1F\x7F]", "",
                                                                                      regex = True).str.replace(r"\s+","",regex = True).tolist())

        sheet_configurations[sheet_name][
            configuration_type
        ] = configured_columns


print("\nConfiguration loaded from:")
print(config_path)

print("\nConfiguration:")
for sheet_name, configuration in sheet_configurations.items():
    print(f"\n{sheet_name}:")
    print(configuration)

# =========================================================
# 3. COLUMN-NAME CLEANING
# =========================================================

def clean_column_names(columns):

    cleaned_columns = []

    for col in columns:

        col = str(col)

        col = re.sub(
            r"[\x00-\x1F\x7F]",
            " ",
            col
        )

        col = re.sub(
            r"\s+",
            " ",
            col
        )

        col = col.strip()
        col = col.replace("-", "_")

        cleaned_columns.append(col)

    return cleaned_columns


# =========================================================
# 4. TEXT CLEANING
# =========================================================

def clean_text(value):

    if pd.isna(value):
        return pd.NA

    if not isinstance(value, str):
        return value

    value = re.sub(
        r"[\x00-\x1F\x7F]",
        " ",
        value
    )

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    value = value.strip()

    missing_markers = {
        "",
        "na",
        "n/a",
        "null",
        "none",
        "nat"
    }

    if value.lower() in missing_markers:
        return pd.NA

    return value


# =========================================================
# 5. MIXED-DATE CONVERSION
# =========================================================

def convert_mixed_dates(date_series):

    converted_dates = pd.Series(
        pd.NaT,
        index=date_series.index,
        dtype="datetime64[ns]"
    )

    text_values = date_series.astype("string")

    # Dates beginning with YYYY-MM-DD or YYYY/MM/DD.
    iso_mask = text_values.str.match(
        r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}",
        na=False
    )

    # Parse ISO dates as year-month-day.
    converted_dates.loc[iso_mask] = pd.to_datetime(
        date_series.loc[iso_mask],
        format="mixed",
        errors="coerce",
        yearfirst=True
    )

    # Parse other formats as day-month-year.
    converted_dates.loc[~iso_mask] = pd.to_datetime(
        date_series.loc[~iso_mask],
        format="mixed",
        errors="coerce",
        dayfirst=True
    )

    return converted_dates


# =========================================================
# 6. CLEAN ONE SHEET
# =========================================================

def clean_dataframe(
    df,
    sheet_name,
    configuration
):

    print("\n" + "=" * 70)
    print(f"CLEANING: {sheet_name}")
    print("=" * 70)

    df = df.copy()


    # -----------------------------------------------------
    # A. Standardize column names
    # -----------------------------------------------------

    df.columns = clean_column_names(
        df.columns
    )

    print("\nColumn names:")
    print(df.columns.tolist())


    # -----------------------------------------------------
    # B. Remove duplicate column names
    # -----------------------------------------------------

    duplicate_columns = (
        df.columns.duplicated().sum()
    )

    print(
        f"Duplicate column names: "
        f"{duplicate_columns}"
    )

    df = df.loc[
        :,
        ~df.columns.duplicated()
    ]


    # -----------------------------------------------------
    # C. Clean text and missing-value markers
    # -----------------------------------------------------

    text_columns = df.select_dtypes(
        include=["object", "string"]
    ).columns

    for col in text_columns:
        df[col] = df[col].map(
            clean_text
        )

    print(
        "Missing-value markers standardized."
    )


    # -----------------------------------------------------
    # D. Remove completely empty rows and columns
    # -----------------------------------------------------

    empty_rows = df.isna().all(
        axis=1
    ).sum()

    empty_columns = df.isna().all(
        axis=0
    ).sum()

    print(
        f"Completely empty rows: {empty_rows}"
    )

    print(
        f"Completely empty columns: {empty_columns}"
    )

    df = df.dropna(
        axis=0,
        how="all"
    )

    df = df.dropna(
        axis=1,
        how="all"
    )

    df = df.reset_index(
        drop=True
    )


    # -----------------------------------------------------
    # E. Remove exact duplicate rows
    # -----------------------------------------------------

    duplicate_rows = df.duplicated().sum()

    print(
        f"Exact duplicate rows: "
        f"{duplicate_rows}"
    )

    df = (
        df.drop_duplicates()
        .reset_index(drop=True)
    )


    # -----------------------------------------------------
    # F. Convert confirmed string columns
    # -----------------------------------------------------

    for col in configuration[
        "string_columns"
    ]:

        if col not in df.columns:
            print(
                f"WARNING: '{col}' was not found."
            )
            continue

        df[col] = df[col].astype(
            "string"
        )


    # -----------------------------------------------------
    # G. Convert numeric columns
    # -----------------------------------------------------

    for col in configuration[
        "numeric_columns"
    ]:

        if col not in df.columns:
            print(
                f"WARNING: '{col}' was not found."
            )
            continue

        original_values = df[col].copy()

        cleaned_numeric_values = (
            original_values
            .astype("string")
            .str.replace(
                ",",
                "",
                regex=False
            )
            .str.replace(
                r"(?i)\bAED\b",
                "",
                regex=True
            )
            .str.strip()
        )

        cleaned_numeric_values = (
            cleaned_numeric_values
            .str.replace(
                r"[^\d.\-]",
                "",
                regex=True
            )
            .replace(
                "",
                pd.NA
            )
        )

        converted_values = pd.to_numeric(
            cleaned_numeric_values,
            errors="coerce"
        )

        failed_mask = (
            original_values.notna()
            & converted_values.isna()
        )

        print(f"\nNumeric column: {col}")

        print(
            f"Invalid numeric values: "
            f"{failed_mask.sum()}"
        )

        if failed_mask.sum() > 0:
            print(
                original_values[failed_mask]
                .value_counts()
                .head(20)
            )

        # Valid values become numeric.
        # Invalid values become missing values.
        df[col] = converted_values


    # -----------------------------------------------------
    # H. Convert date columns
    # -----------------------------------------------------

    for col in configuration[
        "date_columns"
    ]:

        if col not in df.columns:
            print(
                f"WARNING: '{col}' was not found."
            )
            continue

        original_values = df[col].copy()

        original_values = original_values.replace(
            r"^\s*$",
            pd.NaT,
            regex=True
        )

        converted_dates = convert_mixed_dates(
            original_values
        )

        failed_mask = (
            original_values.notna()
            & converted_dates.isna()
        )

        print(f"\nDate column: {col}")

        print(
            f"Invalid dates: "
            f"{failed_mask.sum()}"
        )

        if failed_mask.sum() > 0:
            print(
                original_values[failed_mask]
                .value_counts()
                .head(20)
            )

        # Valid values remain real dates.
        # Empty and invalid dates become pd.NaT.
        df[col] = converted_dates


    # -----------------------------------------------------
    # I. Standardize categories
    # -----------------------------------------------------

    for col in configuration[
        "category_columns"
    ]:

        if col not in df.columns:
            print(
                f"WARNING: '{col}' was not found."
            )
            continue

        df[col] = (
            df[col]
            .astype("string")
            .str.strip()
            .str.title()
        )


    # -----------------------------------------------------
    # J. Final information
    # -----------------------------------------------------

    print("\nFinal data types:")
    print(df.dtypes)

    print("\nMissing values before Excel export:")
    print(df.isna().sum())

    return df


# =========================================================
# 7. READ ALL SHEETS
# =========================================================

excel_file = pd.ExcelFile(
    file_path
)

sheet_list = excel_file.sheet_names

print("Workbook sheets:")
print(sheet_list)


sheets_to_clean = [
    sheet_name
    for sheet_name in sheet_list
    if sheet_name in sheet_configurations
]

all_sheets = pd.read_excel(
    file_path,
    sheet_name=sheets_to_clean
)


# =========================================================
# 8. CLEAN AND SAVE ALL SHEETS
# =========================================================

with pd.ExcelWriter(
    output_path,
    engine="openpyxl"
) as writer:

    for sheet_name, raw_df in all_sheets.items():

        configuration = (
            sheet_configurations[
                sheet_name
            ]
        )

        cleaned_df = clean_dataframe(
            df=raw_df,
            sheet_name=sheet_name,
            configuration=configuration
        )

        cleaned_sheet_name = (
            sheet_name.replace(
                "_Raw",
                "_Cleaned"
            )
        )


        # -------------------------------------------------
        # Create a separate Excel export copy
        # -------------------------------------------------

        export_df = cleaned_df.copy()

        date_columns = configuration[
            "date_columns"
        ]


        # -------------------------------------------------
        # Empty date values → NAT
        # -------------------------------------------------

        for col in date_columns:

            if col not in export_df.columns:
                continue

            # Allows valid dates and "NAT"
            # to exist in the exported column.
            export_df[col] = (
                export_df[col].astype("object")
            )

            export_df.loc[
                export_df[col].isna(),
                col
            ] = "NAT"


        # -------------------------------------------------
        # Other empty values → NA
        # -------------------------------------------------

        non_date_columns = [
            col
            for col in export_df.columns
            if col not in date_columns
        ]

        for col in non_date_columns:

            if export_df[col].isna().sum() > 0:

                export_df[col] = (
                    export_df[col].astype("object")
                )

                export_df.loc[
                    export_df[col].isna(),
                    col
                ] = "NA"


        # -------------------------------------------------
        # Save cleaned sheet
        # -------------------------------------------------

        export_df.to_excel(
            writer,
            sheet_name=cleaned_sheet_name,
            index=False
        )

        worksheet = writer.sheets[
            cleaned_sheet_name
        ]


        # -------------------------------------------------
        # Format valid dates as %d-%b-%y
        # Example: 01-Feb-26
        # -------------------------------------------------

        for date_col in date_columns:

            if date_col not in export_df.columns:
                continue

            column_number = (
                export_df.columns.get_loc(
                    date_col
                )
                + 1
            )

            for row_number in range(
                2,
                len(export_df) + 2
            ):

                cell = worksheet.cell(
                    row=row_number,
                    column=column_number
                )

                if cell.value != "NAT":
                    cell.number_format = (
                        "dd-mmm-yy"
                    )


        # Freeze header row.
        worksheet.freeze_panes = "A2"

        # Apply filters.
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )


        # -------------------------------------------------
        # Adjust column widths
        # -------------------------------------------------

        for column_cells in worksheet.columns:

            column_letter = (
                column_cells[0].column_letter
            )

            maximum_length = 0

            for cell in column_cells:

                if cell.value is None:
                    continue

                maximum_length = max(
                    maximum_length,
                    len(str(cell.value))
                )

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                maximum_length + 2,
                40
            )

        print(
            f"Saved: {cleaned_sheet_name}"
        )


# =========================================================
# 9. FINAL RESULT
# =========================================================

print("\n" + "=" * 70)
print("CLEANING COMPLETED")
print("=" * 70)

print(
    f"Cleaned workbook saved to:\n"
    f"{output_path}"
)
