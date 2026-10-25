import argparse
import re
from pathlib import Path

import pandas as pd


# =========================================================
# 1. CLEANED WORKBOOK PATH
# =========================================================

parser = argparse.ArgumentParser(description="Verify the cleaned sales workbook.")
parser.add_argument("--sample", action="store_true",
                    help="Verify the cleaned portfolio sample.")
args = parser.parse_args()

BASE_DIR = Path(__file__).resolve().parent
output_name = ("portfolio_sample_cleaned.xlsx"
               if args.sample else "online_retail_II_cleaned.xlsx")
file_path = BASE_DIR / "Cleaned" / output_name

# =========================================================
# 2. VERIFICATION CONFIGURATION PATH
# =========================================================

config_path = (
    Path(__file__).resolve().parent
    / "Verification_Configuration.xlsx"
)

if not file_path.exists():
    raise FileNotFoundError(
        f"Cleaned workbook not found:\n{file_path}"
    )

if not config_path.exists():
    raise FileNotFoundError(
        f"Verification configuration not found:\n{config_path}"
    )


# =========================================================
# 3. READ VERIFICATION CONFIGURATION
# =========================================================

column_rules_df = pd.read_excel(
    config_path,
    sheet_name="Column Rules"
)

allowed_values_df = pd.read_excel(
    config_path,
    sheet_name="Allowed Values"
)

date_relationships_df = pd.read_excel(
    config_path,
    sheet_name="Date Relationships"
)


# Clean headers and text stored in the configuration file.
for configuration_df in [
    column_rules_df,
    allowed_values_df,
    date_relationships_df
]:

    configuration_df.columns = (
        configuration_df.columns
        .astype("string")
        .str.strip()
    )

    configuration_text_columns = (
        configuration_df.select_dtypes(
            include=["object", "string"]
        ).columns
    )

    for col in configuration_text_columns:
        configuration_df[col] = (
            configuration_df[col]
            .astype("string")
            .str.strip()
        )


# Confirm that the required configuration headers exist.
required_column_rule_headers = {
    "Sheet",
    "Column",
    "Data Type",
    "Mandatory",
    "Unique",
    "Non Negative"
}

required_allowed_value_headers = {
    "Sheet",
    "Column",
    "Allowed Value"
}

required_date_relationship_headers = {
    "Sheet",
    "Earlier Date",
    "Later Date"
}

configuration_requirements = [
    (
        "Column Rules",
        column_rules_df,
        required_column_rule_headers
    ),
    (
        "Allowed Values",
        allowed_values_df,
        required_allowed_value_headers
    ),
    (
        "Date Relationships",
        date_relationships_df,
        required_date_relationship_headers
    )
]

for config_sheet, config_df, required_headers in (
    configuration_requirements
):

    missing_headers = (
        required_headers - set(config_df.columns)
    )

    if missing_headers:
        raise ValueError(
            f"Missing header(s) in '{config_sheet}': "
            f"{sorted(missing_headers)}"
        )


# Convert the Yes/No Excel rules to True/False.
boolean_rule_columns = [
    "Mandatory",
    "Unique",
    "Non Negative"
]

for col in boolean_rule_columns:

    entered_values = (
        column_rules_df[col]
        .dropna()
        .str.lower()
    )

    invalid_boolean_values = set(
        entered_values.unique()
    ) - {"yes", "no"}

    if invalid_boolean_values:
        raise ValueError(
            f"Invalid value(s) in '{col}': "
            f"{sorted(invalid_boolean_values)}. "
            "Use only Yes or No."
        )

    column_rules_df[col] = (
        column_rules_df[col]
        .str.lower()
        .map({
            "yes": True,
            "no": False
        })
        .fillna(False)
    )


# Build the same sheet_rules dictionary previously written manually.
sheet_rules = {}

for sheet_name, sheet_columns in column_rules_df.groupby(
    "Sheet",
    sort=False
):

    sheet_rules[sheet_name] = {
        "string_columns": [],
        "numeric_columns": [],
        "date_columns": [],
        "mandatory_columns": [],
        "unique_columns": [],
        "non_negative_columns": [],
        "allowed_values": {},
        "date_relationships": []
    }

    for _, row in sheet_columns.iterrows():

        column_name = row["Column"]
        data_type = row["Data Type"].lower()

        data_type_rule = f"{data_type}_columns"

        if data_type_rule not in {
            "string_columns",
            "numeric_columns",
            "date_columns"
        }:
            raise ValueError(
                f"Invalid data type '{data_type}' for "
                f"'{column_name}' in '{sheet_name}'. "
                "Use string, numeric, or date."
            )

        sheet_rules[sheet_name][
            data_type_rule
        ].append(column_name)

        if row["Mandatory"]:
            sheet_rules[sheet_name][
                "mandatory_columns"
            ].append(column_name)

        if row["Unique"]:
            sheet_rules[sheet_name][
                "unique_columns"
            ].append(column_name)

        if row["Non Negative"]:
            sheet_rules[sheet_name][
                "non_negative_columns"
            ].append(column_name)


# Add allowed values to their sheet and column.
for (
    sheet_name,
    column_name
), allowed_group in allowed_values_df.dropna(
    subset=["Sheet", "Column", "Allowed Value"]
).groupby(
    ["Sheet", "Column"],
    sort=False
):

    if sheet_name not in sheet_rules:
        raise ValueError(
            f"Sheet '{sheet_name}' appears in "
            "'Allowed Values' but not in 'Column Rules'."
        )

    if column_name not in sheet_rules[sheet_name][
        "string_columns"
    ]:
        raise ValueError(
            f"Allowed values were configured for "
            f"'{column_name}' in '{sheet_name}', but that "
            "column is not configured as string."
        )

    sheet_rules[sheet_name][
        "allowed_values"
    ][column_name] = (
        allowed_group["Allowed Value"]
        .drop_duplicates()
        .tolist()
    )


# Add earlier-date/later-date relationships.
for _, row in date_relationships_df.dropna(
    subset=["Sheet", "Earlier Date", "Later Date"]
).iterrows():

    sheet_name = row["Sheet"]
    earlier_date = row["Earlier Date"]
    later_date = row["Later Date"]

    if sheet_name not in sheet_rules:
        raise ValueError(
            f"Sheet '{sheet_name}' appears in "
            "'Date Relationships' but not in 'Column Rules'."
        )

    configured_date_columns = sheet_rules[
        sheet_name
    ]["date_columns"]

    if (
        earlier_date not in configured_date_columns
        or later_date not in configured_date_columns
    ):
        raise ValueError(
            f"Date relationship '{earlier_date}' -> "
            f"'{later_date}' in '{sheet_name}' uses a "
            "column that is not configured as date."
        )

    sheet_rules[sheet_name][
        "date_relationships"
    ].append(
        (earlier_date, later_date)
    )


print("Verification rules loaded from:")
print(config_path)


# =========================================================
# 4. VERIFICATION FUNCTION
# =========================================================

def verify_dataframe(df, sheet_name, rules):

    print("\n" + "=" * 70)
    print(f"VERIFICATION REPORT: {sheet_name}")
    print("=" * 70)

    failed_checks = 0


    # -----------------------------------------------------
    # A. Duplicate rows
    # -----------------------------------------------------

    print("\n[A] Verifying duplicate rows...")

    duplicate_rows = df.duplicated().sum()

    if duplicate_rows == 0:
        print("PASS: No duplicate rows detected.")
    else:
        failed_checks += 1
        print(
            f"FAIL: {duplicate_rows} duplicate row(s) detected."
        )


    # -----------------------------------------------------
    # B. Duplicate identifiers
    # -----------------------------------------------------

    print("\n[B] Verifying unique identifiers...")

    failed_unique_columns = []

    if not rules["unique_columns"]:
        print(
            "INFO: No columns are required to contain unique values "
            "according to the verification configuration."
            )

    for col in rules["unique_columns"]:


        if col not in df.columns:
            failed_unique_columns.append(col)
            print(f"FAIL: Unique column '{col}' is missing.")
            continue

        duplicate_identifiers = (
            df[col]
            .dropna()
            .duplicated(keep=False)
        )

        duplicate_count = duplicate_identifiers.sum()

        if duplicate_count == 0:
            print(f"PASS: '{col}' contains unique values.")
        else:
            failed_unique_columns.append(col)

            duplicated_values = (
                df.loc[
                    df[col].notna()
                    & df[col].duplicated(keep=False),
                    col
                ]
                .unique()
                .tolist()
            )

            print(
                f"FAIL: '{col}' contains "
                f"{duplicate_count} duplicated record(s)."
            )
            print(f"Duplicated values: {duplicated_values}")

    if failed_unique_columns:
        failed_checks += 1


    # -----------------------------------------------------
    # C1. String columns
    # -----------------------------------------------------

    print("\n[C1] Verifying string columns...")

    actual_string_columns = df.select_dtypes(
        include=["object", "string"]
    ).columns.tolist()

    failed_string_columns = []

    for col in rules["string_columns"]:

        if col not in df.columns:
            failed_string_columns.append(col)
            print(f"FAIL: String column '{col}' is missing.")

        elif col not in actual_string_columns:
            failed_string_columns.append(col)
            print(
                f"FAIL: '{col}' should be string, "
                f"but its dtype is '{df[col].dtype}'."
            )

        else:
            print(f"PASS: '{col}' is stored as string.")

    if failed_string_columns:
        failed_checks += 1


    # -----------------------------------------------------
    # C2. Numeric columns
    # -----------------------------------------------------

    print("\n[C2] Verifying numeric columns...")

    failed_numeric_columns = []

    for col in rules["numeric_columns"]:

        if col not in df.columns:
            failed_numeric_columns.append(col)
            print(f"FAIL: Numeric column '{col}' is missing.")

        elif not pd.api.types.is_numeric_dtype(df[col]):
            failed_numeric_columns.append(col)
            print(
                f"FAIL: '{col}' should be numeric, "
                f"but its dtype is '{df[col].dtype}'."
            )

        else:
            print(f"PASS: '{col}' is numeric.")
            
        # Business exception counts apply only to valid numeric columns.
        if col not in df.columns or not pd.api.types.is_numeric_dtype(df[col]):
            continue

        if col == "Price":
            zero_prices = df[col].eq(0).sum()
            near_zero_prices = (df[col].gt(0) & df[col].le(0.1)).sum()
            print(f"REVIEW: 'Price' contains {zero_prices} zero price(s).")
            print(
                f"REVIEW: 'Price' contains {near_zero_prices} "
                "near-zero price(s), greater than 0 and up to 0.1 inclusive."
            )
        elif col == "Quantity":
            zero_quantities = df[col].eq(0).sum()
            print(f"REVIEW: 'Quantity' contains {zero_quantities} zero value(s).")


    if failed_numeric_columns:
        failed_checks += 1

    


    # -----------------------------------------------------
    # C3. Date columns
    # -----------------------------------------------------

    print("\n[C3] Verifying date columns...")

    failed_date_columns = []

    for col in rules["date_columns"]:

        if col not in df.columns:
            failed_date_columns.append(col)
            print(f"FAIL: Date column '{col}' is missing.")
            continue

        original_non_null = df[col].notna().sum()

        converted_dates = pd.to_datetime(
            df[col],
            errors="coerce",
            format="mixed"
        )

        valid_dates = converted_dates.notna().sum()
        invalid_dates = original_non_null - valid_dates

        if invalid_dates == 0:
            print(
                f"PASS: Existing values in '{col}' "
                f"are valid dates."
            )
        else:
            failed_date_columns.append(col)

            invalid_values = (
                df.loc[
                    df[col].notna()
                    & converted_dates.isna(),
                    col
                ]
                .value_counts()
                .head(20)
            )

            print(
                f"FAIL: '{col}' contains "
                f"{invalid_dates} invalid date value(s)."
            )
            print("Invalid values:")
            print(invalid_values)

    if failed_date_columns:
        failed_checks += 1


    # -----------------------------------------------------
    # D1. Missing values report
    # -----------------------------------------------------

    print("\n[D1] Reporting missing values...")

    missing_by_column = df.isna().sum()
    columns_with_missing = missing_by_column[
        missing_by_column > 0
    ]

    total_missing_cells = columns_with_missing.sum()

    if total_missing_cells == 0:
        print("PASS: No missing cells detected.")
    else:
        print(
            f"INFO: {total_missing_cells} missing cell(s) "
            f"detected in {len(columns_with_missing)} column(s)."
        )

        for col, count in columns_with_missing.items():
            print(f"  - {col}: {count}")


    # -----------------------------------------------------
    # D2. Mandatory columns
    # -----------------------------------------------------

    print("\n[D2] Verifying mandatory columns...")

    failed_mandatory_columns = []

    for col in rules["mandatory_columns"]:

        if col not in df.columns:
            failed_mandatory_columns.append(col)
            print(f"FAIL: Mandatory column '{col}' is missing.")
            continue

        missing_count = df[col].isna().sum()

        if missing_count == 0:
            print(
                f"PASS: Mandatory column '{col}' "
                f"has no missing values."
            )
        else:
            failed_mandatory_columns.append(col)
            print(
                f"FAIL: Mandatory column '{col}' contains "
                f"{missing_count} missing value(s)."
            )

    if failed_mandatory_columns:
        failed_checks += 1


    # -----------------------------------------------------
    # D3. Completely empty rows and columns
    # -----------------------------------------------------

    print("\n[D3] Verifying empty rows and columns...")

    empty_rows = df.isna().all(axis=1).sum()

    empty_columns = (
        df.columns[df.isna().all(axis=0)]
        .tolist()
    )

    if empty_rows == 0:
        print("PASS: No completely empty rows detected.")
    else:
        failed_checks += 1
        print(
            f"FAIL: {empty_rows} completely empty row(s) detected."
        )

    if not empty_columns:
        print("PASS: No completely empty columns detected.")
    else:
        failed_checks += 1
        print(
            f"FAIL: Completely empty columns detected: "
            f"{empty_columns}"
        )


    # -----------------------------------------------------
    # E. Whitespace
    # -----------------------------------------------------

    print("\n[E] Verifying whitespace...")

    text_columns = df.select_dtypes(
        include=["object", "string"]
    ).columns

    failed_space_columns = []

    for col in text_columns:

        whitespace_mask = (
            df[col]
            .astype("string")
            .str.contains(
                r"^\s|\s$|\s{2,}",
                regex=True,
                na=False
            )
        )

        bad_spaces = whitespace_mask.sum()

        if bad_spaces > 0:
            failed_space_columns.append(col)

            print(
                f"FAIL: '{col}' contains whitespace problems "
                f"in {bad_spaces} row(s)."
            )

    if failed_space_columns:
        failed_checks += 1
    else:
        print("PASS: No whitespace problems detected.")


    # -----------------------------------------------------
    # F. Non-printable characters
    # -----------------------------------------------------

    print("\n[F] Verifying non-printable characters...")

    failed_control_columns = []

    for col in text_columns:

        control_mask = (
            df[col]
            .astype("string")
            .str.contains(
                r"[\x00-\x1F\x7F]",
                regex=True,
                na=False
            )
        )

        control_count = control_mask.sum()

        if control_count > 0:
            failed_control_columns.append(col)

            print(
                f"FAIL: '{col}' contains non-printable "
                f"characters in {control_count} row(s)."
            )

    if failed_control_columns:
        failed_checks += 1
    else:
        print(
            "PASS: No non-printable characters detected."
        )


    # -----------------------------------------------------
    # G1. Duplicate column names
    # -----------------------------------------------------

    print("\n[G1] Verifying duplicate column names...")

    duplicated_headers = (
        df.columns[
            df.columns.duplicated(keep=False)
        ]
        .tolist()
    )

    if not duplicated_headers:
        print("PASS: No duplicate column names detected.")
    else:
        failed_checks += 1
        print(
            f"FAIL: Duplicate column names detected: "
            f"{duplicated_headers}"
        )


    # -----------------------------------------------------
    # G2. Standardized column names
    # -----------------------------------------------------

    print("\n[G2] Verifying standardized column names...")

    failed_header_names = []

    for col in df.columns:

        col_text = str(col)

        has_bad_spacing = bool(
            re.search(
                r"^\s|\s$|\s{2,}",
                col_text
            )
        )

        has_hyphen = "-" in col_text
        is_empty = col_text.strip() == ""

        if has_bad_spacing or has_hyphen or is_empty:
            failed_header_names.append(col)

    if not failed_header_names:
        print("PASS: Column names are standardized.")
    else:
        failed_checks += 1
        print(
            f"FAIL: Non-standard column names detected: "
            f"{failed_header_names}"
        )


    # -----------------------------------------------------
    # H. Negative numeric values
    # -----------------------------------------------------

    print("\n[H] Verifying non-negative numeric values...")

    failed_non_negative_columns = []

    for col in rules["non_negative_columns"]:

        if col not in df.columns:
            continue

        if not pd.api.types.is_numeric_dtype(df[col]):
            print(
                f"SKIPPED: '{col}' is not currently numeric."
            )
            continue

        negative_count = (df[col] < 0).sum()

        if negative_count == 0:
            print(f"PASS: '{col}' contains no negative values.")
        else:
            failed_non_negative_columns.append(col)

            print(
                f"FAIL: '{col}' contains "
                f"{negative_count} negative value(s)."
            )

    if failed_non_negative_columns:
        failed_checks += 1


    # -----------------------------------------------------
    # I. Allowed categorical values
    # -----------------------------------------------------

    print("\n[I] Verifying categorical values...")

    failed_category_columns = []

    if not rules["allowed_values"]:
        print("INFO: No allowed-value rules are configured.")

    for col, allowed_values in rules["allowed_values"].items():


        if col not in df.columns:
            failed_category_columns.append(col)
            print(f"FAIL: Allowed-value column '{col}' is missing.")
            continue

        existing_values = set(
            df[col]
            .dropna()
            .astype(str)
            .unique()
        )

        invalid_values = existing_values - set(allowed_values)

        if not invalid_values:
            print(
                f"PASS: '{col}' contains only allowed values."
            )
        else:
            failed_category_columns.append(col)

            print(
                f"FAIL: '{col}' contains unexpected values:"
            )

            print(sorted(invalid_values))

    if failed_category_columns:
        failed_checks += 1


    # -----------------------------------------------------
    # J. Date relationships
    # -----------------------------------------------------

    print("\n[J] Verifying date relationships...")

    failed_date_relationships = []

    if not rules["date_relationships"]:
        print("INFO: No date relationships are configured.")

    for start_col, end_col in rules["date_relationships"]:

        if (
            start_col not in df.columns
            or end_col not in df.columns
        ):
            failed_date_relationships.append((start_col, end_col))
            print(f"FAIL: Missing date column for '{start_col}' -> '{end_col}'.")
            continue

        start_dates = pd.to_datetime(
            df[start_col],
            errors="coerce",
            format="mixed"
        )

        end_dates = pd.to_datetime(
            df[end_col],
            errors="coerce",
            format="mixed"
        )

        invalid_order_mask = (
            start_dates.notna()
            & end_dates.notna()
            & (end_dates < start_dates)
        )

        invalid_order_count = invalid_order_mask.sum()

        if invalid_order_count == 0:
            print(
                f"PASS: '{end_col}' is not before "
                f"'{start_col}'."
            )
        else:
            failed_date_relationships.append(
                (start_col, end_col)
            )

            print(
                f"FAIL: {invalid_order_count} row(s) have "
                f"'{end_col}' before '{start_col}'."
            )

    if failed_date_relationships:
        failed_checks += 1


    # -----------------------------------------------------
    # FINAL SHEET RESULT
    # -----------------------------------------------------

    print("\n" + "-" * 70)

    if failed_checks == 0:
        print(f"SHEET RESULT: PASS — {sheet_name}")
    else:
        print(
            f"SHEET RESULT: FAIL — {sheet_name} has "
            f"{failed_checks} failed check group(s)."
        )

    return failed_checks


# =========================================================
# 5. READ AND VERIFY ALL SHEETS
# =========================================================

excel_file = pd.ExcelFile(file_path)

all_sheets = pd.read_excel(
    file_path,
    sheet_name=None,
    dtype=object,
    na_values=["NAT"]
)

print("Workbook sheets:")
print(excel_file.sheet_names)

workbook_failed_checks = 0
verified_sheets = 0

for sheet_name, df in all_sheets.items():

    # Preserve Excel text cells such as numeric-looking Invoice IDs.
    # Infer only columns whose underlying values are actually numeric/dates.
    df = df.infer_objects()

    # Skip sheets that do not contain cleaned datasets.
    if sheet_name not in sheet_rules:
        print(
            f"\nSKIPPED: No verification rules configured "
            f"for '{sheet_name}'."
        )
        continue

    failed_checks = verify_dataframe(
        df=df,
        sheet_name=sheet_name,
        rules=sheet_rules[sheet_name]
    )

    workbook_failed_checks += failed_checks
    verified_sheets += 1


# =========================================================
# 6. FINAL WORKBOOK RESULT
# =========================================================

print("\n" + "=" * 70)
print("FINAL WORKBOOK VERIFICATION")
print("=" * 70)

print(f"Sheets verified: {verified_sheets}")
print(f"Total failed check groups: {workbook_failed_checks}")

if verified_sheets == 0:
    print("OVERALL WORKBOOK RESULT: NOT VERIFIED — no matching sheets found.")
    raise SystemExit(1)
elif workbook_failed_checks == 0:
    print("OVERALL WORKBOOK RESULT: PASS")
else:
    print("OVERALL WORKBOOK RESULT: FAIL")
    print("Review the failed checks before approving the workbook.")


if workbook_failed_checks:
    raise SystemExit(1)
