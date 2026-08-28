import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path


# ============================================================
# CSE 4112 - DATA SCIENCE LAB
# LAB-2: DATA CLEANING AND EXPLORATORY DATA ANALYSIS
# ============================================================

INPUT_FILE = "../data_collection/population_data.csv"

CLEANED_FILE = "population_data_cleaned.csv"

EXCEL_FILE = "population_data_cleaned.xlsx"

INSIGHTS_FILE = "cleaning_insights.txt"

GRAPH_DIR = Path("graphs")


# ============================================================
# EXPECTED COLUMNS
# ============================================================

EXPECTED_COLUMNS = [
    "Country",
    "Year",
    "Population",
    "Yearly_Percent_Change",
    "Yearly_Change",
    "Migrants_Net",
    "Median_Age",
    "Fertility_Rate",
    "Density_Per_Km2",
    "Urban_Population_Percent",
    "Urban_Population",
    "Country_Share_of_World_Population",
    "World_Population",
    "Global_Rank"
]


# ============================================================
# PRINT SECTION
# ============================================================

def section(title):

    print()
    print("=" * 75)
    print(title)
    print("=" * 75)


# ============================================================
# LOAD DATA
# ============================================================

def load_data():

    section("STEP 1: LOADING DATA")

    if not Path(INPUT_FILE).exists():

        print(
            f"ERROR: {INPUT_FILE} was not found."
        )

        print(
            "Place Lab-2 in the same folder as "
            "population_data.csv."
        )

        return None

    df = pd.read_csv(
        INPUT_FILE
    )

    print(
        f"Dataset loaded successfully."
    )

    print(
        f"Rows: {len(df):,}"
    )

    print(
        f"Columns: {len(df.columns)}"
    )

    return df


# ============================================================
# INITIAL DATASET INFORMATION
# ============================================================

def initial_analysis(df):

    section("STEP 2: INITIAL DATASET ANALYSIS")

    print("\nColumn names:")

    for column in df.columns:

        print(
            f" - {column}"
        )


    print("\nData types BEFORE cleaning:")

    print(
        df.dtypes
    )


    print(
        "\nMissing values BEFORE cleaning:"
    )

    missing = df.isnull().sum()

    print(
        missing
    )


    print(
        "\nDuplicate rows BEFORE cleaning:",
        df.duplicated().sum()
    )


# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

def clean_column_names(df):

    section("STEP 3: CLEANING COLUMN NAMES")

    df.columns = (
        df.columns
        .str.strip()
        .str.replace(
            " ",
            "_"
        )
    )

    print(
        "Column names standardized."
    )

    return df


# ============================================================
# CLEAN TEXT DATA
# ============================================================

def clean_text_columns(df):

    section("STEP 4: CLEANING TEXT DATA")

    text_columns = [
        "Country"
    ]


    for column in text_columns:

        if column in df.columns:

            df[column] = (
                df[column]
                .astype("string")
                .str.strip()
            )


    # Remove accidental empty country names

    if "Country" in df.columns:

        df["Country"] = (
            df["Country"]
            .replace(
                "",
                pd.NA
            )
        )


    print(
        "Text fields cleaned."
    )

    return df


# ============================================================
# CLEAN NUMERIC DATA
# ============================================================

def clean_numeric_columns(df):

    section("STEP 5: CLEANING NUMERICAL DATA")


    numeric_columns = [

        "Year",

        "Population",

        "Yearly_Percent_Change",

        "Yearly_Change",

        "Migrants_Net",

        "Median_Age",

        "Fertility_Rate",

        "Density_Per_Km2",

        "Urban_Population_Percent",

        "Urban_Population",

        "Country_Share_of_World_Population",

        "World_Population",

        "Global_Rank"
    ]


    for column in numeric_columns:

        if column not in df.columns:

            continue


        # Convert possible strings such as
        # "8,326,979" or "1.38%" into numbers

        df[column] = (

            df[column]

            .astype("string")

            .str.replace(
                ",",
                "",
                regex=False
            )

            .str.replace(
                "%",
                "",
                regex=False
            )

            .str.replace(
                "−",
                "-",
                regex=False
            )

            .str.replace(
                "–",
                "-",
                regex=False
            )

        )


        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        )


    print(
        "Numerical columns converted successfully."
    )

    return df


# ============================================================
# HANDLE INVALID VALUES
# ============================================================

def handle_invalid_values(df):

    section("STEP 6: HANDLING INVALID VALUES")


    # --------------------------------------------------------
    # Year
    # --------------------------------------------------------

    if "Year" in df.columns:

        invalid_years = (
                (df["Year"] < 1900)
                |
                (df["Year"] > 2100)
        )

        print(
            "Invalid Year values:",
            invalid_years.sum()
        )

        df.loc[
            invalid_years,
            "Year"
        ] = np.nan


    # --------------------------------------------------------
    # Population
    # --------------------------------------------------------

    if "Population" in df.columns:

        invalid_population = (
                df["Population"] <= 0
        )

        print(
            "Invalid Population values:",
            invalid_population.sum()
        )

        df.loc[
            invalid_population,
            "Population"
        ] = np.nan


    # --------------------------------------------------------
    # Median Age
    # --------------------------------------------------------

    if "Median_Age" in df.columns:

        invalid_age = (
                (df["Median_Age"] < 0)
                |
                (df["Median_Age"] > 100)
        )

        print(
            "Invalid Median Age values:",
            invalid_age.sum()
        )

        df.loc[
            invalid_age,
            "Median_Age"
        ] = np.nan


    # --------------------------------------------------------
    # Fertility Rate
    # --------------------------------------------------------

    if "Fertility_Rate" in df.columns:

        invalid_fertility = (
                df["Fertility_Rate"] < 0
        )

        print(
            "Invalid Fertility Rate values:",
            invalid_fertility.sum()
        )

        df.loc[
            invalid_fertility,
            "Fertility_Rate"
        ] = np.nan


    # --------------------------------------------------------
    # Density
    # --------------------------------------------------------

    if "Density_Per_Km2" in df.columns:

        invalid_density = (
                df["Density_Per_Km2"] < 0
        )

        print(
            "Invalid Density values:",
            invalid_density.sum()
        )

        df.loc[
            invalid_density,
            "Density_Per_Km2"
        ] = np.nan


    # --------------------------------------------------------
    # Urban population percentage
    # --------------------------------------------------------

    if "Urban_Population_Percent" in df.columns:

        invalid_urban = (

                (df["Urban_Population_Percent"] < 0)

                |

                (df["Urban_Population_Percent"] > 100)

        )

        print(
            "Invalid Urban Population % values:",
            invalid_urban.sum()
        )

        df.loc[
            invalid_urban,
            "Urban_Population_Percent"
        ] = np.nan


    return df


# ============================================================
# REMOVE DUPLICATES
# ============================================================

def remove_duplicates(df):

    section("STEP 7: REMOVING DUPLICATES")


    before = len(df)


    # --------------------------------------------------------
    # Country + Year should uniquely identify a record
    # --------------------------------------------------------

    if (
            "Country" in df.columns
            and
            "Year" in df.columns
    ):

        duplicates = df.duplicated(
            subset=[
                "Country",
                "Year"
            ],
            keep="first"
        )


        print(
            "Duplicate Country-Year records:",
            duplicates.sum()
        )


        df = df[
            ~duplicates
        ].copy()


    else:

        df = df.drop_duplicates()


    after = len(df)


    print(
        f"Rows before: {before:,}"
    )

    print(
        f"Rows after: {after:,}"
    )

    print(
        f"Duplicates removed: {before - after:,}"
    )


    return df


# ============================================================
# HANDLE MISSING VALUES
# ============================================================

def handle_missing_values(df):

    section("STEP 8: HANDLING MISSING VALUES")


    print(
        "Missing values BEFORE:"
    )

    print(
        df.isnull().sum()
    )


    # --------------------------------------------------------
    # Numeric columns
    #
    # We use median imputation because population data may
    # contain legitimate extreme values.
    # --------------------------------------------------------

    numeric_columns = df.select_dtypes(
        include=np.number
    ).columns


    for column in numeric_columns:

        missing_count = (
            df[column]
            .isnull()
            .sum()
        )


        if missing_count == 0:

            continue


        median_value = (
            df[column]
            .median()
        )


        if pd.notna(
                median_value
        ):

            df[column] = (
                df[column]
                .fillna(
                    median_value
                )
            )


        print(
            f"{column}: "
            f"{missing_count} missing "
            f"values handled."
        )


    # --------------------------------------------------------
    # Country
    # --------------------------------------------------------

    if "Country" in df.columns:

        df["Country"] = (
            df["Country"]
            .fillna(
                "Unknown"
            )
        )


    print(
        "\nMissing values AFTER:"
    )

    print(
        df.isnull().sum()
    )


    return df


# ============================================================
# FIX DATA TYPES
# ============================================================

def fix_data_types(df):

    section("STEP 9: FIXING DATA TYPES")


    if "Year" in df.columns:

        df["Year"] = (
            df["Year"]
            .round()
            .astype(int)
        )


    if "Global_Rank" in df.columns:

        df["Global_Rank"] = (
            df["Global_Rank"]
            .round()
            .astype(int)
        )


    integer_columns = [

        "Population",

        "Yearly_Change",

        "Migrants_Net",

        "Density_Per_Km2",

        "Urban_Population",

        "World_Population"

    ]


    for column in integer_columns:

        if column in df.columns:

            df[column] = (
                df[column]
                .round()
                .astype("int64")
            )


    print(
        "\nData types AFTER cleaning:"
    )

    print(
        df.dtypes
    )


    return df


# ============================================================
# SORT DATA
# ============================================================

def sort_data(df):

    section("STEP 10: SORTING DATA")


    if (
            "Country" in df.columns
            and
            "Year" in df.columns
    ):

        df = df.sort_values(
            by=[
                "Country",
                "Year"
            ]
        )


        df = df.reset_index(
            drop=True
        )


    print(
        "Dataset sorted by Country and Year."
    )


    return df


# ============================================================
# GENERATE CLEANING INSIGHTS
# ============================================================

def generate_cleaning_insights(
        original_df,
        cleaned_df
):

    section("STEP 11: CLEANING INSIGHTS")


    insights = []


    # --------------------------------------------------------
    # Dataset size
    # --------------------------------------------------------

    original_rows = len(
        original_df
    )

    cleaned_rows = len(
        cleaned_df
    )


    insights.append(
        "DATA CLEANING INSIGHTS"
    )

    insights.append(
        "=" * 60
    )


    insights.append(
        f"Original number of records: "
        f"{original_rows:,}"
    )


    insights.append(
        f"Final number of records: "
        f"{cleaned_rows:,}"
    )


    # --------------------------------------------------------
    # Missing values
    # --------------------------------------------------------

    original_missing = int(
        original_df.isnull()
        .sum()
        .sum()
    )


    final_missing = int(
        cleaned_df.isnull()
        .sum()
        .sum()
    )


    insights.append(
        f"Missing values before cleaning: "
        f"{original_missing:,}"
    )


    insights.append(
        f"Missing values after cleaning: "
        f"{final_missing:,}"
    )


    # --------------------------------------------------------
    # Duplicates
    # --------------------------------------------------------

    duplicate_count = (
        original_df.duplicated(
            subset=[
                "Country",
                "Year"
            ]
        ).sum()
    )


    insights.append(
        f"Duplicate Country-Year records removed: "
        f"{duplicate_count:,}"
    )


    # --------------------------------------------------------
    # Country count
    # --------------------------------------------------------

    if "Country" in cleaned_df.columns:

        country_count = (
            cleaned_df["Country"]
            .nunique()
        )

        insights.append(
            f"Unique countries: "
            f"{country_count:,}"
        )


    # --------------------------------------------------------
    # Year range
    # --------------------------------------------------------

    if "Year" in cleaned_df.columns:

        min_year = (
            cleaned_df["Year"]
            .min()
        )

        max_year = (
            cleaned_df["Year"]
            .max()
        )

        insights.append(
            f"Year range: "
            f"{min_year} - {max_year}"
        )


    # --------------------------------------------------------
    # Population statistics
    # --------------------------------------------------------

    if "Population" in cleaned_df.columns:

        largest = (
            cleaned_df.loc[
                cleaned_df["Population"].idxmax()
            ]
        )


        smallest = (
            cleaned_df.loc[
                cleaned_df["Population"].idxmin()
            ]
        )


        insights.append(
            f"Largest population record: "
            f"{largest['Country']} "
            f"({largest['Population']:,})"
        )


        insights.append(
            f"Smallest population record: "
            f"{smallest['Country']} "
            f"({smallest['Population']:,})"
        )


    # --------------------------------------------------------
    # Cleaning operations
    # --------------------------------------------------------

    insights.append("")
    insights.append(
        "CLEANING OPERATIONS PERFORMED"
    )

    insights.append(
        "-" * 60
    )


    insights.append(
        "1. Standardized column names."
    )

    insights.append(
        "2. Removed unnecessary whitespace "
        "from text values."
    )

    insights.append(
        "3. Converted numerical columns "
        "to appropriate numeric data types."
    )

    insights.append(
        "4. Removed commas from large numerical "
        "values."
    )

    insights.append(
        "5. Removed percentage symbols from "
        "percentage fields."
    )

    insights.append(
        "6. Normalized negative values."
    )

    insights.append(
        "7. Identified and removed duplicate "
        "Country-Year records."
    )

    insights.append(
        "8. Detected invalid numerical values."
    )

    insights.append(
        "9. Handled missing numerical values "
        "using median imputation."
    )

    insights.append(
        "10. Sorted the final dataset by "
        "Country and Year."
    )


    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    with open(
            INSIGHTS_FILE,
            "w",
            encoding="utf-8"
    ) as file:

        for line in insights:

            file.write(
                line + "\n"
            )


    print(
        f"Cleaning insights saved to: "
        f"{INSIGHTS_FILE}"
    )


# ============================================================
# BASIC EDA
# ============================================================

def perform_eda(df):

    section("STEP 12: EXPLORATORY DATA ANALYSIS")


    # --------------------------------------------------------
    # Descriptive statistics
    # --------------------------------------------------------

    print(
        "\nDESCRIPTIVE STATISTICS"
    )


    print(
        df.describe()
    )


    # ========================================================
    # GRAPH DIRECTORY
    # ========================================================

    GRAPH_DIR.mkdir(
        exist_ok=True
    )


    # ========================================================
    # GRAPH 1
    # TOP 10 POPULATED COUNTRY-YEAR RECORDS
    # ========================================================

    if "Population" in df.columns:

        top10 = (
            df.nlargest(
                10,
                "Population"
            )
        )


        plt.figure(
            figsize=(12, 6)
        )


        plt.bar(
            top10["Country"],
            top10["Population"]
        )


        plt.title(
            "Top 10 Population Records"
        )


        plt.xlabel(
            "Country"
        )


        plt.ylabel(
            "Population"
        )


        plt.xticks(
            rotation=45,
            ha="right"
        )


        plt.tight_layout()


        plt.savefig(
            GRAPH_DIR /
            "top_10_population.png",
            dpi=300
        )


        plt.close()


    # ========================================================
    # GRAPH 2
    # POPULATION GROWTH OVER TIME
    # ========================================================

    if (
            "Year" in df.columns
            and
            "Population" in df.columns
    ):

        yearly_population = (
            df.groupby(
                "Year"
            )["Population"]
            .sum()
        )


        plt.figure(
            figsize=(12, 6)
        )


        plt.plot(
            yearly_population.index,
            yearly_population.values
        )


        plt.title(
            "Total Population by Year"
        )


        plt.xlabel(
            "Year"
        )


        plt.ylabel(
            "Total Population"
        )


        plt.grid(
            True,
            alpha=0.3
        )


        plt.tight_layout()


        plt.savefig(
            GRAPH_DIR /
            "population_over_time.png",
            dpi=300
        )


        plt.close()


    # ========================================================
    # GRAPH 3
    # FERTILITY RATE DISTRIBUTION
    # ========================================================

    if "Fertility_Rate" in df.columns:

        plt.figure(
            figsize=(10, 6)
        )


        plt.hist(
            df["Fertility_Rate"],
            bins=20
        )


        plt.title(
            "Distribution of Fertility Rate"
        )


        plt.xlabel(
            "Fertility Rate"
        )


        plt.ylabel(
            "Frequency"
        )


        plt.tight_layout()


        plt.savefig(
            GRAPH_DIR /
            "fertility_rate_distribution.png",
            dpi=300
        )


        plt.close()


    # ========================================================
    # GRAPH 4
    # URBANIZATION
    # ========================================================

    if (
            "Urban_Population_Percent" in df.columns
            and
            "Year" in df.columns
    ):

        urban_by_year = (
            df.groupby(
                "Year"
            )["Urban_Population_Percent"]
            .mean()
        )


        plt.figure(
            figsize=(12, 6)
        )


        plt.plot(
            urban_by_year.index,
            urban_by_year.values
        )


        plt.title(
            "Average Urban Population Percentage"
        )


        plt.xlabel(
            "Year"
        )


        plt.ylabel(
            "Urban Population (%)"
        )


        plt.grid(
            True,
            alpha=0.3
        )


        plt.tight_layout()


        plt.savefig(
            GRAPH_DIR /
            "urban_population_trend.png",
            dpi=300
        )


        plt.close()


    print(
        f"Graphs saved inside: {GRAPH_DIR}"
    )


# ============================================================
# SAVE CLEANED CSV
# ============================================================

def save_csv(df):

    section("STEP 13: SAVING CLEANED CSV")


    df.to_csv(
        CLEANED_FILE,
        index=False,
        encoding="utf-8-sig"
    )


    print(
        f"Cleaned CSV created: "
        f"{CLEANED_FILE}"
    )


# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(df):

    section("STEP 14: SAVING EXCEL FILE")


    with pd.ExcelWriter(
            EXCEL_FILE,
            engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Cleaned Data",
            index=False
        )


        # --------------------------------------------
        # Statistics sheet
        # --------------------------------------------

        statistics = df.describe(
            include="all"
        ).transpose()


        statistics.to_excel(
            writer,
            sheet_name="Statistics"
        )


    print(
        f"Excel file created: "
        f"{EXCEL_FILE}"
    )


# ============================================================
# FINAL SUMMARY
# ============================================================

def final_summary(df):

    section("LAB-2 FINAL RESULT")


    print(
        f"Final records: "
        f"{len(df):,}"
    )


    print(
        f"Final columns: "
        f"{len(df.columns)}"
    )


    print(
        f"Unique countries: "
        f"{df['Country'].nunique():,}"
    )


    print(
        f"Year range: "
        f"{df['Year'].min()} - "
        f"{df['Year'].max()}"
    )


    print(
        "\nRemaining missing values:"
    )

    print(
        df.isnull().sum()
    )


    print()
    print(
        "OUTPUT FILES"
    )

    print(
        f"1. {CLEANED_FILE}"
    )

    print(
        f"2. {EXCEL_FILE}"
    )

    print(
        f"3. {INSIGHTS_FILE}"
    )

    print(
        f"4. {GRAPH_DIR}/"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 75)
    print("CSE 4112 - DATA SCIENCE LAB")
    print("LAB-2: DATA CLEANING + EDA")
    print("=" * 75)


    # --------------------------------------------------------
    # Load
    # --------------------------------------------------------

    df = load_data()


    if df is None:

        return


    # Keep original copy
    original_df = df.copy(
        deep=True
    )


    # --------------------------------------------------------
    # Initial analysis
    # --------------------------------------------------------

    initial_analysis(
        df
    )


    # --------------------------------------------------------
    # Cleaning pipeline
    # --------------------------------------------------------

    df = clean_column_names(
        df
    )

    df = clean_text_columns(
        df
    )

    df = clean_numeric_columns(
        df
    )

    df = handle_invalid_values(
        df
    )

    df = remove_duplicates(
        df
    )

    df = handle_missing_values(
        df
    )

    df = fix_data_types(
        df
    )

    df = sort_data(
        df
    )


    # --------------------------------------------------------
    # Cleaning insights
    # --------------------------------------------------------

    generate_cleaning_insights(
        original_df,
        df
    )


    # --------------------------------------------------------
    # EDA
    # --------------------------------------------------------

    perform_eda(
        df
    )


    # --------------------------------------------------------
    # Save files
    # --------------------------------------------------------

    save_csv(
        df
    )

    save_excel(
        df
    )


    # --------------------------------------------------------
    # Final summary
    # --------------------------------------------------------

    final_summary(
        df
    )


    print()
    print("=" * 75)
    print("LAB-2 COMPLETED SUCCESSFULLY")
    print("=" * 75)


# ============================================================
# RUN PROGRAM
# ============================================================

if __name__ == "__main__":

    main()