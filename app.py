from __future__ import annotations

import io
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill

st.set_page_config(
    page_title="Hospital Data Integration Studio",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

PROJECT_TITLE = "Hospital Multi-Format Data Integration, EDA and Preprocessing"

# Dataset files that are shipped with the app (put them in a folder called "data"
# next to this app.py file in your GitHub repo).
DATA_DIR = Path(__file__).parent / "data"
SAMPLE_FILES = [
    ("CSV", "hospital_source.csv"),
    ("JSON", "hospital_source.json"),
    ("XML", "hospital_source.xml"),
    ("HTML", "hospital_source.html"),
    ("KML", "hospital_source.kml"),
    ("KMZ", "hospital_source.kmz"),
]

COMMON_COLUMNS = [
    "Hospital_Name", "Address", "City", "ZIP_Code", "Phone", "Website",
    "County_Status", "Longitude", "Latitude", "Location", "Data_Year",
    "Hospital_Type", "Bed_Count", "Quality_Score", "Record_Status",
]
NUMERIC_COLUMNS = ["Longitude", "Latitude", "Data_Year", "Bed_Count", "Quality_Score"]
TEXT_COLUMNS = [
    "Hospital_Name", "Address", "City", "Phone", "Website",
    "County_Status", "Hospital_Type", "Record_Status",
]

# ---- EDA settings (taken from the notebook) ----
# Tokens that really mean "missing"; they are converted to NaN before EDA (notebook CELL 20).
MISSING_TOKENS = ["", " ", "NA", "N/A", "NULL", "null", "None", "NONE",
                  "UNKNOWN", "Unknown", "unknown"]
CATEGORICAL_COLUMNS = ["County_Status", "Hospital_Type", "Record_Status", "Source_Format"]
OUTLIER_COLUMNS = ["Bed_Count", "Quality_Score", "Latitude", "Longitude"]

# Every record gets a Source_Record_ID (1, 2, 3 ... inside each file). Because that ID is
# unique, NO two rows can ever be "exact duplicates" if the ID is included in the comparison.
# True  -> duplicates are detected on the real hospital fields (+ Source_Format), ignoring the ID.
# False -> exactly the notebook behaviour (ID included, so duplicates are never found).
IGNORE_ID_COLUMNS_FOR_DUPLICATES = True

FIELD_DESCRIPTIONS = {
    "Hospital_Name": ("Text", "Name of the hospital / facility"),
    "Address": ("Text", "Street address"),
    "City": ("Text", "City / town / municipality"),
    "ZIP_Code": ("Text (5 digits)", "Postal code"),
    "Phone": ("Text (10 digits)", "Contact phone number"),
    "Website": ("Text", "Hospital website / web link"),
    "County_Status": ("Category", "In-County or Out-of-County flag"),
    "Longitude": ("Numeric", "Longitude coordinate (-180 to 180)"),
    "Latitude": ("Numeric", "Latitude coordinate (-90 to 90)"),
    "Location": ("Text", "Raw coordinate string (mainly from KML/KMZ)"),
    "Data_Year": ("Numeric", "Reporting / data year"),
    "Hospital_Type": ("Category", "General Hospital, Emergency Center, Specialty Hospital, etc."),
    "Bed_Count": ("Numeric", "Number of beds (must be >= 0)"),
    "Quality_Score": ("Numeric", "Quality / inspection score (0 to 100)"),
    "Record_Status": ("Text", "Status flag of the record (HTML source)"),
}

# These mappings reproduce the source notebook's field mappings.
SCHEMA_MAPPING = {
    "CSV": {
        "Hospital_Name": "Hospital_Name", "Address": "Address", "City": "City",
        "ZIP_Code": "ZIP_Code", "Phone": "Phone", "Website": "Website",
        "County_Status": "County_Status", "Longitude": "Longitude",
        "Latitude": "Latitude", "Location": "Location", "Data_Year": "Data_Year",
        "Hospital_Type": "Hospital_Type", "Bed_Count": "Bed_Count",
        "Quality_Score": "Inspection_Score", "Record_Status": None,
    },
    "JSON": {
        "Hospital_Name": "hospital_name", "Address": "street_address", "City": "city",
        "ZIP_Code": "postal_code", "Phone": "telephone", "Website": "web",
        "County_Status": "county_status", "Longitude": "lon", "Latitude": "lat",
        "Location": None, "Data_Year": "year", "Hospital_Type": "hospital_type",
        "Bed_Count": "beds", "Quality_Score": "rating", "Record_Status": None,
    },
    "XML": {
        "Hospital_Name": "HospitalName", "Address": "Street", "City": "Town",
        "ZIP_Code": "PostalCode", "Phone": "ContactPhone", "Website": "Website",
        "County_Status": "CountyStatus", "Longitude": "Longitude",
        "Latitude": "Latitude", "Location": None, "Data_Year": "Year",
        "Hospital_Type": "Type", "Bed_Count": "Beds", "Quality_Score": "QualityScore",
        "Record_Status": None,
    },
    "HTML": {
        "Hospital_Name": "Facility Name", "Address": "Street Address",
        "City": "Municipality", "ZIP_Code": "Postal", "Phone": "Contact",
        "Website": "Web Link", "County_Status": "County Flag", "Longitude": "Lon",
        "Latitude": "Lat", "Location": None, "Data_Year": "Reporting Year",
        "Hospital_Type": "Facility Category", "Bed_Count": "Beds Available",
        "Quality_Score": None, "Record_Status": "Record Status",
    },
    "KML": {
        "Hospital_Name": "name", "Address": None, "City": None, "ZIP_Code": None,
        "Phone": None, "Website": None, "County_Status": None, "Longitude": None,
        "Latitude": None, "Location": "coordinates", "Data_Year": None,
        "Hospital_Type": None, "Bed_Count": None, "Quality_Score": None,
        "Record_Status": None,
    },
    "KMZ": {
        "Hospital_Name": "name", "Address": None, "City": None, "ZIP_Code": None,
        "Phone": None, "Website": None, "County_Status": None, "Longitude": None,
        "Latitude": None, "Location": "coordinates", "Data_Year": None,
        "Hospital_Type": None, "Bed_Count": None, "Quality_Score": None,
        "Record_Status": None,
    },
}

ALIASES = {
    "Hospital_Name": ["hospital name", "facility name", "name", "hospital_name", "hospitalname"],
    "Address": ["address", "street", "street address", "street_address"],
    "City": ["city", "town", "municipality"],
    "ZIP_Code": ["zip", "zip code", "postal", "postal code", "postal_code", "zipcode"],
    "Phone": ["phone", "telephone", "contact", "contactphone", "contact phone"],
    "Website": ["website", "web", "web link", "url"],
    "County_Status": ["county status", "county_status", "county flag"],
    "Longitude": ["longitude", "lon", "lng"],
    "Latitude": ["latitude", "lat"],
    "Location": ["location", "coordinates"],
    "Data_Year": ["data year", "year", "reporting year"],
    "Hospital_Type": ["hospital type", "hospital_type", "type", "facility category"],
    "Bed_Count": ["bed count", "beds", "bed_count", "beds available"],
    "Quality_Score": ["quality score", "qualityscore", "rating", "inspection score", "inspection_score"],
    "Record_Status": ["record status", "record_status"],
}


# ---------------------------------------------------------------------
# Core processing functions
# ---------------------------------------------------------------------
def norm_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def detect_format(filename: str) -> str:
    ext = Path(filename).suffix.lower()
    mapping = {".csv": "CSV", ".json": "JSON", ".xml": "XML", ".html": "HTML",
               ".htm": "HTML", ".kml": "KML", ".kmz": "KMZ"}
    if ext not in mapping:
        raise ValueError(f"Unsupported file extension: {ext or '(none)'}.")
    return mapping[ext]


def parse_kml_bytes(data: bytes) -> pd.DataFrame:
    root = ET.fromstring(data)
    rows = []
    for placemark in root.iter():
        if placemark.tag.split("}")[-1] != "Placemark":
            continue
        row = {}
        for child in placemark.iter():
            tag = child.tag.split("}")[-1]
            if tag == "name" and child.text and "name" not in row:
                row["name"] = child.text.strip()
            elif tag == "description" and child.text:
                row["description"] = child.text.strip()
            elif tag == "coordinates" and child.text:
                row["coordinates"] = child.text.strip()
        row.setdefault("coordinates", np.nan)
        rows.append(row)
    return pd.DataFrame(rows)


def parse_upload(uploaded_file, source_format: str) -> pd.DataFrame:
    data = uploaded_file.getvalue()
    if source_format == "CSV":
        return pd.read_csv(io.BytesIO(data))
    if source_format == "JSON":
        obj = json.loads(data.decode("utf-8-sig"))
        if isinstance(obj, dict):
            candidate = next((v for v in obj.values() if isinstance(v, list)), None)
            obj = candidate if candidate is not None else [obj]
        return pd.json_normalize(obj)
    if source_format == "XML":
        root = ET.fromstring(data)
        records = []
        for element in root.iter():
            if element.tag.split("}")[-1].lower() in {"hospital", "record", "facility"}:
                record = {}
                for child in list(element):
                    record[child.tag.split("}")[-1]] = child.text
                if record:
                    records.append(record)
        if not records:
            for element in list(root):
                record = {child.tag.split("}")[-1]: child.text for child in list(element)}
                if record:
                    records.append(record)
        return pd.DataFrame(records)
    if source_format == "HTML":
        tables = pd.read_html(io.BytesIO(data))
        if not tables:
            raise ValueError("No HTML table was found in this file.")
        return tables[0].copy()
    if source_format == "KML":
        return parse_kml_bytes(data)
    if source_format == "KMZ":
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            kml_names = [n for n in archive.namelist()
                         if n.lower().endswith(".kml") and not n.startswith("__MACOSX/")]
            if not kml_names:
                raise ValueError("No .kml file was found inside this KMZ archive.")
            return parse_kml_bytes(archive.read(kml_names[0]))
    raise ValueError(f"Unsupported format: {source_format}")


class LocalFile:
    """Wraps a file on disk so it behaves like a Streamlit uploaded file."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.name = self.path.name

    def getvalue(self) -> bytes:
        return self.path.read_bytes()


def sample_files() -> list[LocalFile]:
    return [LocalFile(DATA_DIR / name) for _, name in SAMPLE_FILES if (DATA_DIR / name).exists()]


@st.cache_data(show_spinner=False)
def load_sample_raw(filename: str) -> pd.DataFrame:
    f = LocalFile(DATA_DIR / filename)
    return parse_upload(f, detect_format(filename))


def standardize_schema(df: pd.DataFrame, source_format: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    mapping = SCHEMA_MAPPING[source_format]
    lookup = {norm_name(c): c for c in df.columns}
    standardized = pd.DataFrame(index=df.index)
    mapping_rows = []
    for target in COMMON_COLUMNS:
        source = mapping.get(target)
        chosen = source if source is not None and source in df.columns else None
        if chosen is None:
            target_norm = norm_name(target)
            if target_norm in lookup:
                chosen = lookup[target_norm]
            else:
                for alias in ALIASES.get(target, []):
                    if norm_name(alias) in lookup:
                        chosen = lookup[norm_name(alias)]
                        break
        standardized[target] = df[chosen] if chosen is not None else np.nan
        if chosen is None:
            transformation = "Not available in source"
        elif chosen == target:
            transformation = "Same field name"
        else:
            transformation = "Renamed / mapped"
        mapping_rows.append({
            "Source_Format": source_format, "Common_Field": target,
            "Source_Field": chosen if chosen is not None else "Not available",
            "Mapping_Status": "Mapped" if chosen is not None else "Missing in source",
            "Transformation": transformation,
        })
    standardized["Source_Format"] = source_format
    standardized["Source_Record_ID"] = range(1, len(standardized) + 1)
    return standardized, pd.DataFrame(mapping_rows)


def make_missing_report(df: pd.DataFrame) -> pd.DataFrame:
    report = pd.DataFrame({
        "Column": df.columns,
        "Missing_Count": [int(df[c].isna().sum()) for c in df.columns],
    })
    report["Missing_Percentage"] = (
        report["Missing_Count"] / max(len(df), 1) * 100
    ).round(2)
    return report.sort_values("Missing_Count", ascending=False)


def duplicate_columns(df: pd.DataFrame) -> list[str]:
    """Columns used to decide whether two rows are exact duplicates."""
    if IGNORE_ID_COLUMNS_FOR_DUPLICATES:
        return [c for c in df.columns if c != "Source_Record_ID"]
    return list(df.columns)


# ---------------------------------------------------------------------
# EDA (follows the notebook, CELL 20 - CELL 38)
# ---------------------------------------------------------------------
def standardize_missing_tokens(df: pd.DataFrame) -> pd.DataFrame:
    """Notebook CELL 20: turn tokens such as 'NA', 'N/A', 'unknown' into real missing values."""
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == "object":
            out[col] = out[col].where(~out[col].isin(MISSING_TOKENS), np.nan)
    return out


def iqr_outliers(df: pd.DataFrame, col: str) -> tuple[dict, pd.DataFrame]:
    """Notebook CELL 33: IQR rule. Returns a summary row and the outlier records."""
    vals = pd.to_numeric(df[col], errors="coerce")
    q1, q3 = vals.quantile(0.25), vals.quantile(0.75)
    iqr = q3 - q1
    lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    mask = (vals < lower) | (vals > upper)
    rows = df.loc[mask, ["Hospital_Name", "Source_Format", col]]
    summary = {"Column": col, "Q1": q1, "Q3": q3, "IQR": iqr,
               "Lower_Limit": lower, "Upper_Limit": upper, "Outlier_Count": int(mask.sum())}
    return summary, rows


def run_eda(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """
    Runs the notebook's EDA on the combined, schema-standardized data (before cleaning).
    Returns (eda_text_frame, eda_numeric_frame, eda_results).
      eda_text_frame    -> missing tokens standardized, numeric columns still as read
      eda_numeric_frame -> same, with numeric columns converted (this is what preprocessing starts from)
    """
    eda_text = standardize_missing_tokens(raw)

    # --- shape, data types, missing values (CELL 21-24) ---
    dtype_report = pd.DataFrame({"Column": eda_text.columns,
                                 "Data_Type": eda_text.dtypes.astype(str).values})
    missing_report = make_missing_report(eda_text)
    missing_by_source = (eda_text.isna().groupby(eda_text["Source_Format"]).sum().sum(axis=1)
                         .reset_index(name="Total_Missing_Cells"))
    missing_heatmap = (eda_text[COMMON_COLUMNS].isna().groupby(eda_text["Source_Format"]).mean() * 100).round(1)

    # --- duplicates (CELL 25) ---
    dup_cols = duplicate_columns(eda_text)
    exact_dups = int(eda_text.duplicated(subset=dup_cols).sum())
    potential_mask = eda_text.duplicated(subset=["Hospital_Name", "Address", "City"], keep=False)
    potential_dups = int(eda_text.duplicated(subset=["Hospital_Name", "Address", "City"]).sum())
    dup_rows = eda_text[eda_text.duplicated(subset=dup_cols, keep=False)].sort_values(
        [c for c in ["Hospital_Name", "Source_Format"] if c in eda_text.columns])
    potential_rows = eda_text[potential_mask].sort_values("Hospital_Name")

    # --- unique values (CELL 26) ---
    unique_report = pd.DataFrame({
        "Column": eda_text.columns,
        "Unique_Values": [int(eda_text[c].nunique(dropna=True)) for c in eda_text.columns],
    })

    # --- categorical values (CELL 27) ---
    categorical = {}
    for col in CATEGORICAL_COLUMNS:
        counts = (eda_text[col].fillna("Missing").astype(str).value_counts()
                  .rename_axis(col).reset_index(name="Count"))
        categorical[col] = counts

    # --- numeric conversion + descriptive statistics (CELL 28-29) ---
    eda_num = eda_text.copy()
    for col in NUMERIC_COLUMNS:
        eda_num[col] = pd.to_numeric(eda_num[col], errors="coerce")
    numeric_summary = eda_num[NUMERIC_COLUMNS].describe().T

    # --- invalid values (CELL 30-32) ---
    invalid_beds = eda_num[eda_num["Bed_Count"] < 0]
    invalid_quality = eda_num[(eda_num["Quality_Score"] < 0) | (eda_num["Quality_Score"] > 100)]
    invalid_coords = eda_num[(eda_num["Longitude"] < -180) | (eda_num["Longitude"] > 180) |
                             (eda_num["Latitude"] < -90) | (eda_num["Latitude"] > 90)]

    # --- IQR outliers for 4 numeric columns (CELL 33 + outlier cells) ---
    outlier_rows, outlier_records = [], {}
    for col in OUTLIER_COLUMNS:
        summary, rows = iqr_outliers(eda_num, col)
        outlier_rows.append(summary)
        outlier_records[col] = rows
    outlier_analysis = pd.DataFrame(outlier_rows)

    # --- EDA issue log (CELL 38; coordinate and potential-duplicate issues added) ---
    n = max(len(eda_num), 1)
    issues = []
    for _, row in missing_report.iterrows():
        if row["Missing_Count"] > 0:
            issues.append({"Issue_Type": "Missing Values", "Column": row["Column"],
                           "Count": int(row["Missing_Count"]), "Percentage": float(row["Missing_Percentage"]),
                           "Action_Required": "Handle during preprocessing"})
    if len(invalid_beds) > 0:
        issues.append({"Issue_Type": "Invalid Value", "Column": "Bed_Count", "Count": len(invalid_beds),
                       "Percentage": round(len(invalid_beds) / n * 100, 2),
                       "Action_Required": "Convert invalid negative values to NA"})
    if len(invalid_quality) > 0:
        issues.append({"Issue_Type": "Invalid Value", "Column": "Quality_Score", "Count": len(invalid_quality),
                       "Percentage": round(len(invalid_quality) / n * 100, 2),
                       "Action_Required": "Convert values outside 0-100 to NA"})
    if len(invalid_coords) > 0:
        issues.append({"Issue_Type": "Invalid Value", "Column": "Latitude / Longitude", "Count": len(invalid_coords),
                       "Percentage": round(len(invalid_coords) / n * 100, 2),
                       "Action_Required": "Convert out-of-range coordinates to NA"})
    if exact_dups > 0:
        issues.append({"Issue_Type": "Duplicates", "Column": "Entire Record", "Count": exact_dups,
                       "Percentage": round(exact_dups / n * 100, 2),
                       "Action_Required": "Remove exact duplicates"})
    if potential_dups > 0:
        issues.append({"Issue_Type": "Potential Duplicates", "Column": "Hospital_Name + Address + City",
                       "Count": potential_dups, "Percentage": round(potential_dups / n * 100, 2),
                       "Action_Required": "Review (not removed automatically)"})
    for col in OUTLIER_COLUMNS:
        cnt = int(outlier_analysis.loc[outlier_analysis["Column"] == col, "Outlier_Count"].iloc[0])
        if cnt > 0:
            issues.append({"Issue_Type": "IQR Outliers", "Column": col, "Count": cnt,
                           "Percentage": round(cnt / n * 100, 2),
                           "Action_Required": "Review (reported, not deleted)"})
    issue_log = pd.DataFrame(issues, columns=["Issue_Type", "Column", "Count", "Percentage", "Action_Required"])

    eda = {
        "rows": len(eda_text), "columns": len(eda_text.columns),
        "dtypes": dtype_report, "missing": missing_report,
        "missing_by_source": missing_by_source, "missing_heatmap": missing_heatmap,
        "exact_dups": exact_dups, "potential_dups": potential_dups,
        "dup_rows": dup_rows, "potential_rows": potential_rows,
        "unique": unique_report, "categorical": categorical,
        "numeric_summary": numeric_summary,
        "invalid_beds": invalid_beds, "invalid_quality": invalid_quality, "invalid_coords": invalid_coords,
        "outliers": outlier_analysis, "outlier_records": outlier_records,
        "issue_log": issue_log,
    }
    return eda_text, eda_num, eda


# ---------------------------------------------------------------------
# Preprocessing
# ---------------------------------------------------------------------
def preprocess_data(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict]:
    clean = raw.copy()
    before_rows = len(clean)
    before_missing = int(clean.isna().sum().sum())
    before_duplicates = int(clean.duplicated(subset=duplicate_columns(clean)).sum())
    actions = []

    for col in TEXT_COLUMNS:
        clean[col] = clean[col].astype("string").str.strip()
    actions.append({"Step": "Text standardization", "Action": "Trimmed whitespace from text fields"})

    county_mapping = {
        "in county": "In-County", "in-county": "In-County",
        "out county": "Out-of-County", "out-of-county": "Out-of-County",
    }
    clean["County_Status"] = clean["County_Status"].apply(
        lambda x: county_mapping.get(str(x).strip().lower(), x) if pd.notna(x) else x
    )
    actions.append({"Step": "County category standardization", "Action": "Normalized county status labels"})

    type_mapping = {
        "general": "General Hospital", "emergency": "Emergency Center",
        "specialty": "Specialty Hospital",
    }
    clean["Hospital_Type"] = clean["Hospital_Type"].apply(
        lambda x: type_mapping.get(str(x).strip().lower(), x) if pd.notna(x) else x
    )
    actions.append({"Step": "Hospital type standardization", "Action": "Normalized common hospital type labels"})

    for col in NUMERIC_COLUMNS:
        clean[col] = pd.to_numeric(clean[col], errors="coerce")
    actions.append({"Step": "Numeric conversion", "Action": "Converted numeric fields; invalid text becomes missing"})

    invalid_beds = int((clean["Bed_Count"] < 0).sum())
    clean.loc[clean["Bed_Count"] < 0, "Bed_Count"] = np.nan
    invalid_quality = int(((clean["Quality_Score"] < 0) | (clean["Quality_Score"] > 100)).sum())
    clean.loc[(clean["Quality_Score"] < 0) | (clean["Quality_Score"] > 100), "Quality_Score"] = np.nan

    invalid_lon = (clean["Longitude"] < -180) | (clean["Longitude"] > 180)
    invalid_lat = (clean["Latitude"] < -90) | (clean["Latitude"] > 90)
    coordinate_errors = int(invalid_lon.sum() + invalid_lat.sum())
    clean.loc[invalid_lon, "Longitude"] = np.nan
    clean.loc[invalid_lat, "Latitude"] = np.nan

    clean["ZIP_Code"] = clean["ZIP_Code"].astype("string").str.strip().str.replace(r"\.0$", "", regex=True)
    zip_valid = clean["ZIP_Code"].str.match(r"^\d{5}$", na=False)
    invalid_zip = int((~zip_valid & clean["ZIP_Code"].notna()).sum())
    clean.loc[~zip_valid, "ZIP_Code"] = pd.NA

    def clean_phone(value):
        if pd.isna(value):
            return pd.NA
        digits = re.sub(r"\D", "", str(value))
        return digits if len(digits) == 10 else pd.NA

    clean["Phone"] = clean["Phone"].apply(clean_phone)
    clean["Website"] = clean["Website"].astype("string").str.strip().replace("", pd.NA)
    clean["Hospital_Name"] = clean["Hospital_Name"].astype("string").str.strip().str.replace(
        "National Institue", "National Institute", regex=False
    )

    rows_before_dedup = len(clean)
    clean = clean.drop_duplicates(subset=duplicate_columns(clean)).reset_index(drop=True)
    duplicates_removed = rows_before_dedup - len(clean)

    imputation_rows = []
    for col in ["Bed_Count", "Quality_Score"]:
        missing_before = int(clean[col].isna().sum())
        median = clean[col].median()
        if pd.notna(median):
            clean[col] = clean[col].fillna(median)
        imputation_rows.append({
            "Column": col, "Missing_Before": missing_before,
            "Imputation_Method": "Median",
            "Imputation_Value": median if pd.notna(median) else "No median available",
            "Missing_After": int(clean[col].isna().sum()),
        })
    imputation_report = pd.DataFrame(imputation_rows)

    validation = []
    checks = [
        ("Negative Bed Count", int((clean["Bed_Count"] < 0).sum())),
        ("Quality Score 0-100", int(((clean["Quality_Score"] < 0) | (clean["Quality_Score"] > 100)).sum())),
        ("Latitude Range", int(((clean["Latitude"] < -90) | (clean["Latitude"] > 90)).sum())),
        ("Longitude Range", int(((clean["Longitude"] < -180) | (clean["Longitude"] > 180)).sum())),
        ("Duplicate Rows", int(clean.duplicated(subset=duplicate_columns(clean)).sum())),
        ("Invalid ZIP format", int((clean["ZIP_Code"].notna() & ~clean["ZIP_Code"].str.match(r"^\d{5}$", na=False)).sum())),
        ("Invalid phone format", int((clean["Phone"].notna() & ~clean["Phone"].str.match(r"^\d{10}$", na=False)).sum())),
    ]
    for name, count in checks:
        validation.append({"Validation_Check": name, "Invalid_Count": count,
                           "Status": "PASS" if count == 0 else "FAIL"})
    validation_report = pd.DataFrame(validation)

    actions.extend([
        {"Step": "Invalid values", "Action": f"Invalid beds: {invalid_beds}; invalid quality scores: {invalid_quality}; invalid coordinate values: {coordinate_errors} set to missing"},
        {"Step": "ZIP validation", "Action": f"{invalid_zip} invalid ZIP values set to missing"},
        {"Step": "Phone standardization", "Action": "Kept only 10-digit phone numbers; others set to missing"},
        {"Step": "Hospital name standardization", "Action": "Applied notebook spelling correction for 'National Institue'"},
        {"Step": "Duplicate removal", "Action": f"Removed {duplicates_removed} exact duplicate rows"},
        {"Step": "Median imputation", "Action": "Filled missing Bed_Count and Quality_Score with column medians when available"},
    ])
    stats = {
        "rows_before": before_rows, "rows_after": len(clean),
        "columns": len(clean.columns), "missing_before": before_missing,
        "missing_after": int(clean.isna().sum().sum()),
        "duplicates_before": before_duplicates, "duplicates_removed": duplicates_removed,
        "invalid_beds": invalid_beds, "invalid_quality": invalid_quality,
        "invalid_coordinates": coordinate_errors, "invalid_zip": invalid_zip,
    }
    return clean, validation_report, imputation_report, {"stats": stats, "actions": pd.DataFrame(actions)}


def create_excel_report(raw, clean, source_summary, schema_report, missing_report,
                        validation_report, imputation_report, actions, before_after,
                        dtype_report, outlier_report, eda_issue_log, final_source_summary) -> bytes:
    final_display = clean.copy().fillna("NA")
    out = io.BytesIO()
    sheets = {
        "Raw_Combined_Data": raw, "Final_Clean_Data": final_display,
        "Source_Summary": source_summary, "Schema_Mapping": schema_report,
        "Data_Types": dtype_report, "Missing_Values": missing_report,
        "EDA_Issue_Log": eda_issue_log,
        "Outlier_Analysis": outlier_report, "Before_After": before_after,
        "Preprocessing_Log": actions, "Validation_Report": validation_report,
        "Imputation_Report": imputation_report,
        "Final_Source_Summary": final_source_summary,
    }
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="D9EAF7")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for column_cells in ws.columns:
                max_len = max((len(str(c.value)) for c in column_cells if c.value is not None), default=10)
                ws.column_dimensions[column_cells[0].column_letter].width = min(max_len + 2, 35)
            for row in ws.iter_rows():
                for cell in row:
                    if str(cell.value).strip().upper() == "NA":
                        cell.fill = PatternFill("solid", fgColor="FFF2CC")
    return out.getvalue()


def process_files(uploads) -> bool:
    """Parse, standardize, run EDA, preprocess and store results in session state."""
    parsed, standardized, mapping_frames, errors = [], [], [], []
    progress = st.progress(0, text="Reading files…")
    for i, file in enumerate(uploads):
        try:
            fmt = detect_format(file.name)
            df = parse_upload(file, fmt)
            if df.empty:
                raise ValueError("The file was read, but no records were found.")
            std, mapping = standardize_schema(df, fmt)
            parsed.append({"filename": file.name, "format": fmt, "raw": df, "standardized": std})
            standardized.append(std)
            mapping_frames.append(mapping.assign(Source_File=file.name))
        except Exception as exc:
            errors.append({"File": file.name, "Error": str(exc)})
        progress.progress((i + 1) / len(uploads), text=f"Processed {i + 1} of {len(uploads)} file(s)")
    progress.empty()
    if errors:
        st.warning("Some files could not be processed. Review the details below; successfully read files are still available.")
        st.dataframe(pd.DataFrame(errors), use_container_width=True, hide_index=True)
    if not standardized:
        st.error("No files were processed successfully. Check file formats and structure.")
        return False

    combined = pd.concat(standardized, ignore_index=True)
    mapping_report = pd.concat(mapping_frames, ignore_index=True)
    source_summary = pd.DataFrame([
        {"File": p["filename"], "Source_Format": p["format"],
         "Input_Rows": len(p["raw"]), "Input_Columns": len(p["raw"].columns),
         "Mapped_Fields": int((mapping_report.loc[mapping_report["Source_File"] == p["filename"], "Mapping_Status"] == "Mapped").sum())}
        for p in parsed
    ])

    # EDA first (on the combined data, before any cleaning), then preprocessing starts from its output.
    eda_text, eda_num, eda = run_eda(combined)
    clean, validation_report, imputation_report, metadata = preprocess_data(eda_num)
    stats = metadata["stats"]

    final_source_summary = (clean.groupby("Source_Format").size().reset_index(name="Final_Record_Count"))
    before_after = pd.DataFrame([
        {"Metric": "Rows", "Before_Preprocessing": len(eda_num), "After_Preprocessing": len(clean)},
        {"Metric": "Missing cells", "Before_Preprocessing": stats["missing_before"], "After_Preprocessing": stats["missing_after"]},
        {"Metric": "Exact duplicate rows", "Before_Preprocessing": stats["duplicates_before"],
         "After_Preprocessing": int(clean.duplicated(subset=duplicate_columns(clean)).sum())},
        {"Metric": "Invalid negative bed counts", "Before_Preprocessing": len(eda["invalid_beds"]),
         "After_Preprocessing": int((clean["Bed_Count"] < 0).sum())},
        {"Metric": "Invalid quality scores", "Before_Preprocessing": len(eda["invalid_quality"]),
         "After_Preprocessing": int(((clean["Quality_Score"] < 0) | (clean["Quality_Score"] > 100)).sum())},
    ])
    st.session_state["hospital_result"] = {
        "parsed": parsed, "raw": eda_num, "clean": clean, "mapping": mapping_report,
        "source_summary": source_summary, "missing": eda["missing"], "dtypes": eda["dtypes"],
        "validation": validation_report, "imputation": imputation_report,
        "actions": metadata["actions"], "before_after": before_after,
        "outliers": eda["outliers"], "errors": errors,
        "eda": eda, "final_source_summary": final_source_summary,
    }
    return True


# ---------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------
st.markdown("""
<style>
.block-container {padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1500px;}
.hero {padding: 1.3rem 1.5rem; border-radius: 18px; background: linear-gradient(120deg,#102a43,#176b87); color: white; margin-bottom: 1.2rem;}
.hero h1 {font-size: 2.1rem; margin: 0 0 .4rem 0;}
.hero p {margin: 0; opacity: .9;}
.card {border: 1px solid rgba(128,128,128,.25); border-radius: 14px; padding: 1rem 1.2rem; margin-bottom: 1rem;}
.card h4 {margin-top: 0;}
div[data-testid="stMetric"] {border: 1px solid rgba(128,128,128,.22); padding: 14px; border-radius: 14px;}

/* ---------- About page cards ---------- */
.feature-card {border: 1px solid rgba(128,128,128,.25); border-radius: 14px; padding: 1rem 1.1rem; height: 100%;
               border-top: 4px solid #176b87; margin-bottom: .8rem;}
.feature-card .icon {font-size: 1.6rem; margin-bottom: .3rem;}
.feature-card h5 {margin: 0 0 .35rem 0; font-size: 1.02rem;}
.feature-card p {margin: 0; font-size: .92rem; opacity: .85; line-height: 1.5;}
.stage {display: flex; gap: .9rem; align-items: flex-start; padding: .8rem 1rem; margin-bottom: .6rem;
        border: 1px solid rgba(128,128,128,.22); border-radius: 12px;}
.stage .num {min-width: 34px; height: 34px; border-radius: 50%; background: linear-gradient(135deg,#102a43,#176b87);
             color: #fff; display: flex; align-items: center; justify-content: center; font-weight: 700;}
.stage .body b {display: block; margin-bottom: .15rem;}
.stage .body span {font-size: .92rem; opacity: .85; line-height: 1.5;}
.badge {display: inline-block; padding: .18rem .65rem; margin: .15rem .2rem .15rem 0; border-radius: 999px;
        background: rgba(23,107,135,.14); border: 1px solid rgba(23,107,135,.35); font-size: .85rem;}

/* ---------- Sidebar menu ---------- */
section[data-testid="stSidebar"] {background: linear-gradient(180deg,#0b1f33 0%,#102a43 55%,#143d57 100%);}
section[data-testid="stSidebar"] * {color: #e8f1f7;}
section[data-testid="stSidebar"] hr {border-color: rgba(255,255,255,.14);}
.side-brand {display: flex; align-items: center; gap: .7rem; padding: .4rem .2rem 1rem .2rem;}
.side-brand .logo {width: 46px; height: 46px; border-radius: 14px; font-size: 1.6rem;
                   background: linear-gradient(135deg,#1fa2c7,#176b87); display: flex;
                   align-items: center; justify-content: center; box-shadow: 0 4px 14px rgba(0,0,0,.35);}
.side-brand .title {font-weight: 700; font-size: 1.05rem; line-height: 1.15;}
.side-brand .sub {font-size: .74rem; opacity: .7; letter-spacing: .04em;}
.side-label {font-size: .72rem; letter-spacing: .12em; text-transform: uppercase; opacity: .6; margin: .2rem 0 .5rem .3rem;}

section[data-testid="stSidebar"] div[role="radiogroup"] {gap: .35rem;}
section[data-testid="stSidebar"] div[role="radiogroup"] > label {
    width: 100%; padding: .65rem .85rem; border-radius: 12px; cursor: pointer;
    border: 1px solid transparent; transition: all .18s ease;
    background: rgba(255,255,255,.04);
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
    background: rgba(255,255,255,.12); transform: translateX(3px);
}
/* hide the round radio dot */
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {display: none;}
section[data-testid="stSidebar"] div[role="radiogroup"] > label p {font-size: .98rem; font-weight: 500; margin: 0;}
/* active page */
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
    background: linear-gradient(90deg,#1fa2c7,#176b87);
    border-color: rgba(255,255,255,.35);
    box-shadow: 0 6px 18px rgba(0,0,0,.35);
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p {font-weight: 700; color: #fff;}

.status-pill {display: flex; align-items: center; gap: .55rem; padding: .6rem .8rem; border-radius: 12px;
              background: rgba(255,255,255,.06); border: 1px solid rgba(255,255,255,.12); font-size: .85rem;}
.dot {width: 10px; height: 10px; border-radius: 50%; display: inline-block;}
.dot.on {background: #3ddc97; box-shadow: 0 0 8px #3ddc97;}
.dot.off {background: #f5b942; box-shadow: 0 0 8px #f5b942;}
.side-foot {font-size: .75rem; opacity: .6; margin-top: .8rem; line-height: 1.5;}
</style>
<div class="hero">
<h1>🏥 Hospital Data Integration Studio</h1>
<p>Integrate heterogeneous hospital data, explore data quality, preprocess records, validate, and export.</p>
</div>
""", unsafe_allow_html=True)


def require_result():
    result = st.session_state.get("hospital_result")
    # A result saved by an older version of the app has no "eda" key, so ask for a re-run.
    if not result or "eda" not in result:
        st.info("No results yet. Open **▶️ Implementation** in the sidebar, run the bundled dataset files "
                "(or upload your own) and the results will appear here.")
        return None
    return result


# ---------------------------------------------------------------------
# 1. ABOUT PROJECT
# ---------------------------------------------------------------------
def page_about_project():
    st.header("ℹ️ About the Project")

    st.markdown(f"""
<div class="card">
<h4>{PROJECT_TITLE}</h4>
Hospital information is published by many different agencies, and each one tends to use its own file format,
its own column names, its own units and its own level of quality. One agency may share a spreadsheet, another a
JSON feed, a third an XML export, and a map service may only provide placemarks with coordinates. Before any
meaningful analysis can happen, all of this has to be brought together into one consistent, trustworthy dataset.<br><br>
This project is a single web application that <b>integrates</b> hospital records from <b>six file formats</b>
(CSV, JSON, XML, HTML, KML and KMZ) into one common <b>15-field schema</b>. It then performs
<b>Exploratory Data Analysis (EDA)</b> to expose real data-quality problems, applies a documented
<b>preprocessing</b> pipeline to fix them, runs automated <b>validation</b> checks on the result, and finally
<b>exports</b> analysis-ready data together with a complete, formatted audit report.
</div>
""", unsafe_allow_html=True)

    # ---------------- Problem & motivation ----------------
    st.subheader("🧩 Problem statement and motivation")
    p1, p2 = st.columns(2)
    with p1:
        st.markdown("**The problem**")
        st.markdown(
            "- The same hospital attribute has different names in different files "
            "(e.g. `Inspection_Score`, `rating`, `QualityScore` all mean the quality score)\n"
            "- Formats range from flat tables to nested documents, web tables and geographic files\n"
            "- Values are inconsistent: ZIP codes with stray decimals, phone numbers with symbols, "
            "category labels written several ways, negative bed counts, impossible coordinates\n"
            "- Some sources simply do not contain certain fields (for example KML only has a name and coordinates)\n"
            "- Duplicates and spelling mistakes creep in when records are merged"
        )
    with p2:
        st.markdown("**Why it matters**")
        st.markdown(
            "- Analysts spend most of their time cleaning data rather than analysing it\n"
            "- Decisions about capacity, access and quality of care depend on correct beds, locations and scores\n"
            "- Dirty data silently produces misleading charts, wrong averages and failed map plots\n"
            "- A repeatable, documented pipeline is far more reliable than one-off manual fixes\n"
            "- Every step must be transparent so the results can be trusted and reviewed"
        )

    # ---------------- Objectives ----------------
    st.subheader("🎯 Objectives")
    objectives = [
        ("🔗", "Integrate", "Read six different file formats and combine them into one standard schema without losing the source of each record."),
        ("🔍", "Profile", "Measure missing values, data types, duplicates and outliers so data-quality problems become visible."),
        ("🧹", "Clean", "Standardize text, categories, numbers, ZIP codes and phone numbers using clearly documented rules."),
        ("✅", "Validate", "Run automated PASS / FAIL checks after cleaning to prove the rules actually worked."),
        ("📦", "Deliver", "Export cleaned data, a schema-mapping file and a complete multi-sheet Excel report."),
        ("🔁", "Reproduce", "Make the whole process repeatable, so any new set of files can be processed with one click."),
    ]
    for row_start in range(0, len(objectives), 3):
        cols = st.columns(3)
        for col, (icon, title, text) in zip(cols, objectives[row_start:row_start + 3]):
            with col:
                st.markdown(
                    f'<div class="feature-card"><div class="icon">{icon}</div><h5>{title}</h5><p>{text}</p></div>',
                    unsafe_allow_html=True,
                )

    # ---------------- Pipeline ----------------
    st.subheader("⚙️ How the pipeline works")
    stages = [
        ("Parse", "Each file is read with a format-specific reader: tables for CSV, normalized records for JSON, "
                  "repeated elements for XML, the first table for HTML, placemarks for KML, and the embedded KML for zipped KMZ files."),
        ("Schema mapping", "Source columns are mapped to the 15 common fields using explicit mappings first, then normalized "
                           "names and common aliases. Fields a source does not provide stay missing and are never invented."),
        ("Combine", "All standardized records are stacked into one table, tagged with their source format and a record ID "
                    "so every row can be traced back to its file."),
        ("Explore (EDA)", "Missing values, data types, duplicates, unique values, category counts, invalid values, descriptive "
                          "statistics and IQR outliers are analysed on the combined data before cleaning, so real problems are visible."),
        ("Preprocess", "Twelve documented rules trim text, normalize categories, convert numbers, nullify invalid values, "
                       "validate ZIP and phone formats, fix spelling, remove duplicates and impute medians."),
        ("Validate", "Seven automated checks confirm there are no negative beds, out-of-range scores or coordinates, "
                     "duplicates, or badly formatted ZIP codes and phone numbers."),
        ("Export", "Download the cleaned CSV, the schema mapping, the validation report, or one formatted Excel workbook "
                   "containing every table produced by the pipeline."),
    ]
    for i, (title, text) in enumerate(stages, start=1):
        st.markdown(
            f'<div class="stage"><div class="num">{i}</div><div class="body"><b>{title}</b><span>{text}</span></div></div>',
            unsafe_allow_html=True,
        )

    # ---------------- Features ----------------
    st.subheader("✨ Key features")
    f1, f2, f3 = st.columns(3)
    features = [
        (f1, "📂 Multi-format ingestion", "CSV, JSON, XML, HTML, KML and KMZ in a single run. Use the bundled dataset or upload your own files."),
        (f2, "🧭 Transparent mapping", "A full source-to-common field mapping table shows exactly which source column fed each field."),
        (f3, "📊 Interactive EDA", "Plotly charts and summary tables for missing data, duplicates, categories, outliers and distributions."),
        (f1, "🛡️ Safe cleaning", "Invalid values become missing instead of being guessed. Descriptive fields like phone and address are never fabricated."),
        (f2, "🧾 Full audit trail", "An EDA issue log, preprocessing log, imputation report and before/after comparison document every change made."),
        (f3, "⬇️ One-click export", "Cleaned CSV, mapping CSV, validation CSV and a formatted multi-sheet Excel report."),
    ]
    for col, title, text in features:
        with col:
            st.markdown(
                f'<div class="feature-card"><h5>{title}</h5><p>{text}</p></div>',
                unsafe_allow_html=True,
            )

    # ---------------- Challenges ----------------
    st.subheader("🛠️ Data challenges handled")
    st.dataframe(pd.DataFrame([
        {"Challenge": "Different column names", "Example": "Inspection_Score / rating / QualityScore", "How it is handled": "Explicit mapping, then normalized-name and alias matching"},
        {"Challenge": "Inconsistent categories", "Example": "“in county”, “In-County”", "How it is handled": "Label dictionaries convert variants to one standard value"},
        {"Challenge": "Impossible values", "Example": "Negative beds, score of 140, latitude of 95", "How it is handled": "Range checks set invalid values to missing"},
        {"Challenge": "Malformed identifiers", "Example": "ZIP “12345.0”, phone “(555) 123-4567”", "How it is handled": "Strip suffixes and symbols, keep only valid 5-digit / 10-digit values"},
        {"Challenge": "Fields absent in a source", "Example": "KML has no address or bed count", "How it is handled": "Left as missing, reported in the mapping table"},
        {"Challenge": "Duplicates and typos", "Example": "“National Institue”", "How it is handled": "Spelling correction and exact duplicate removal"},
        {"Challenge": "Missing numeric data", "Example": "Blank Bed_Count or Quality_Score", "How it is handled": "Median imputation, with every filled value logged"},
    ]), use_container_width=True, hide_index=True)

    # ---------------- Outcomes ----------------
    st.subheader("🏁 Expected outcomes")
    o1, o2 = st.columns(2)
    with o1:
        st.markdown("**Deliverables**")
        st.markdown(
            "- One unified, schema-standardized hospital dataset\n"
            "- A cleaned dataset ready for analysis, modelling or mapping\n"
            "- A schema-mapping report and a validation report\n"
            "- A multi-sheet Excel workbook documenting the entire process"
        )
    with o2:
        st.markdown("**What the result enables**")
        st.markdown(
            "- Comparing hospitals across regions, types and sizes\n"
            "- Plotting facilities on a map using validated coordinates\n"
            "- Studying capacity (beds) and quality scores with reliable numbers\n"
            "- Reusing the same pipeline on new files with no extra effort"
        )

    # ---------------- Tech ----------------
    st.subheader("🧰 Technologies used")
    st.dataframe(pd.DataFrame([
        {"Technology": "Python", "Purpose": "Core language for the whole pipeline"},
        {"Technology": "Streamlit", "Purpose": "Interactive multi-page web application"},
        {"Technology": "Pandas", "Purpose": "Parsing, merging, cleaning and profiling tabular data"},
        {"Technology": "NumPy", "Purpose": "Numeric operations and missing-value handling"},
        {"Technology": "Plotly", "Purpose": "Interactive charts for EDA and overview pages"},
        {"Technology": "OpenPyXL", "Purpose": "Building the formatted Excel report"},
        {"Technology": "lxml / html5lib", "Purpose": "Reading HTML tables and XML-based formats"},
    ]), use_container_width=True, hide_index=True)

    st.markdown("**Supported input formats**")
    st.markdown("".join(f'<span class="badge">{fmt}</span>' for fmt, _ in SAMPLE_FILES), unsafe_allow_html=True)


# ---------------------------------------------------------------------
# 2. ARCHITECTURE  (new page, sits above "About Dataset")
# ---------------------------------------------------------------------
ARCHITECTURE_DOT = """
digraph G {
  rankdir=LR; bgcolor="transparent"; nodesep=0.25; ranksep=0.55; compound=true;
  node [shape=box, style="rounded,filled", fontname="Helvetica", fontsize=11,
        color="#176b87", fillcolor="#e8f4f8", fontcolor="#102a43", margin="0.16,0.08"];
  edge [color="#176b87", arrowsize=0.8];

  subgraph cluster_src {
    label="1  Data sources"; style="rounded,filled"; fillcolor="#f4f9fb"; color="#9bc4d3"; fontname="Helvetica"; fontcolor="#102a43";
    CSV; JSON; XML; HTML; KML; KMZ;
  }
  P  [label="2  Format-specific parsers\\nparse_upload()", fillcolor="#d5ebf3"];
  S  [label="3  Schema mapping\\nstandardize_schema()\\n15 common fields", fillcolor="#d5ebf3"];
  C  [label="4  Combine\\npd.concat + Source_Format\\n+ Source_Record_ID", fillcolor="#d5ebf3"];
  E  [label="5  EDA\\nrun_eda()\\nmissing · duplicates · outliers", fillcolor="#fff2cc", color="#c9a227"];
  PP [label="6  Preprocessing\\npreprocess_data()\\n12 rules", fillcolor="#d5ebf3"];
  V  [label="7  Validation\\n7 PASS / FAIL checks", fillcolor="#dff3e4", color="#3b8f5a"];
  X  [label="8  Export\\nExcel · CSV", fillcolor="#dff3e4", color="#3b8f5a"];

  CSV -> P; JSON -> P; XML -> P; HTML -> P; KML -> P; KMZ -> P;
  P -> S -> C -> E -> PP -> V -> X;
  E -> PP [style=invis];
}
"""


def page_architecture():
    st.header("🏗️ Architecture")
    st.write("The application is a single Streamlit file organised in four layers. Data flows left to right: "
             "raw files are parsed, mapped to one schema, combined, explored (EDA), cleaned, validated and exported.")

    st.subheader("End-to-end data flow")
    st.graphviz_chart(ARCHITECTURE_DOT, use_container_width=True)
    st.caption("Yellow = analysis stage (EDA, runs before any cleaning). Green = output / quality stages.")

    st.subheader("Application layers")
    layers = [
        ("🖥️", "Presentation layer", "Streamlit sidebar with six pages: About Project, Architecture, About Dataset, "
                                     "Preprocessing, EDA and Implementation. Custom CSS gives the cards and navigation their look."),
        ("⚙️", "Processing layer", "Plain Python functions in app.py: parsers, schema mapper, EDA, preprocessing, validation "
                                  "and Excel report builder. They contain no Streamlit code, so they are easy to test."),
        ("🗄️", "Data layer", "Six source files in the data/ folder (or user uploads). Everything is processed in memory; "
                            "results are kept in st.session_state so every page can read them after a run."),
        ("📤", "Output layer", "Cleaned CSV, standardized CSV, schema-mapping CSV, validation CSV and a multi-sheet Excel "
                              "workbook with NA cells highlighted."),
    ]
    cols = st.columns(4)
    for col, (icon, title, text) in zip(cols, layers):
        with col:
            st.markdown(f'<div class="feature-card"><div class="icon">{icon}</div><h5>{title}</h5><p>{text}</p></div>',
                        unsafe_allow_html=True)

    st.subheader("Component reference")
    st.dataframe(pd.DataFrame([
        {"Stage": "Parse", "Function": "parse_upload(), parse_kml_bytes()", "Input": "CSV / JSON / XML / HTML / KML / KMZ file", "Output": "One raw DataFrame per file"},
        {"Stage": "Schema mapping", "Function": "standardize_schema()", "Input": "Raw DataFrame + SCHEMA_MAPPING / ALIASES", "Output": "15-field DataFrame + mapping report"},
        {"Stage": "Combine", "Function": "process_files()", "Input": "Six standardized DataFrames", "Output": "One combined table tagged with Source_Format and Source_Record_ID"},
        {"Stage": "EDA", "Function": "run_eda()", "Input": "Combined table", "Output": "Missing, duplicate, unique, category, invalid-value, outlier tables and the EDA issue log"},
        {"Stage": "Preprocess", "Function": "preprocess_data()", "Input": "EDA output (numeric-converted)", "Output": "Clean table, imputation report, preprocessing log"},
        {"Stage": "Validate", "Function": "preprocess_data() (validation block)", "Input": "Clean table", "Output": "Seven PASS / FAIL checks"},
        {"Stage": "Export", "Function": "create_excel_report()", "Input": "All report tables", "Output": "Formatted 13-sheet Excel workbook"},
    ]), use_container_width=True, hide_index=True)

    left, right = st.columns(2)
    with left:
        st.subheader("Project structure")
        st.code(
            "Intelligent-Data-Contract-Framework/\n"
            "├── app.py                  # Streamlit application (all pages + pipeline)\n"
            "├── requirements.txt        # streamlit, pandas, plotly, openpyxl, lxml, html5lib\n"
            "├── README.md\n"
            "└── data/\n"
            "    ├── hospital_source.csv\n"
            "    ├── hospital_source.json\n"
            "    ├── hospital_source.xml\n"
            "    ├── hospital_source.html\n"
            "    ├── hospital_source.kml\n"
            "    └── hospital_source.kmz",
            language="text",
        )
    with right:
        st.subheader("Design decisions")
        st.markdown(
            "- **EDA before cleaning** so the issue log shows the real problems, and preprocessing starts from the EDA output\n"
            "- **Never invent descriptive data**: invalid phone, ZIP, address values become missing\n"
            "- **Median imputation only** for Bed_Count and Quality_Score\n"
            "- **Source traceability**: every row keeps Source_Format and Source_Record_ID\n"
            "- **In-memory processing**: uploaded files are not written to disk\n"
            "- **Reusable**: the same pipeline runs on the bundled data or any new set of files"
        )


# ---------------------------------------------------------------------
# 3. ABOUT DATASET
# ---------------------------------------------------------------------
def page_about_dataset():
    st.header("📚 About the Dataset")
    st.write("The project uses hospital records supplied as **six data files**, one per format. "
             "Each file describes the same kind of facilities but uses its own structure and column names. "
             "Every source is converted into one **common 15-field schema**.")

    st.subheader("Dataset files")
    available = {name: (DATA_DIR / name).exists() for _, name in SAMPLE_FILES}
    if not all(available.values()):
        missing = [n for n, ok in available.items() if not ok]
        st.warning("Some dataset files were not found in the `data/` folder: " + ", ".join(missing))

    rows = []
    for fmt, name in SAMPLE_FILES:
        path = DATA_DIR / name
        if not path.exists():
            continue
        try:
            df = load_sample_raw(name)
            rows.append({"Format": fmt, "File": name, "Size (KB)": round(path.stat().st_size / 1024, 1),
                         "Records": len(df), "Fields in source": len(df.columns)})
        except Exception as exc:
            rows.append({"Format": fmt, "File": name, "Size (KB)": round(path.stat().st_size / 1024, 1),
                         "Records": "error", "Fields in source": str(exc)})
    if rows:
        summary = pd.DataFrame(rows)
        st.dataframe(summary, use_container_width=True, hide_index=True)
        numeric_rows = summary[pd.to_numeric(summary["Records"], errors="coerce").notna()]
        if not numeric_rows.empty:
            st.caption(f"Total records across all files: {int(pd.to_numeric(numeric_rows['Records']).sum()):,}")

    st.subheader("Preview each file")
    tabs = st.tabs([fmt for fmt, _ in SAMPLE_FILES])
    notes = {
        "CSV": "Tabular rows and columns. Quality score comes from `Inspection_Score`.",
        "JSON": "List of records. Short keys such as `lat`, `lon`, `beds`, `rating`.",
        "XML": "Repeated `<Hospital>` elements; child elements become fields.",
        "HTML": "First HTML table in the file. Has `Record Status` but no quality score.",
        "KML": "Map placemarks. Only the name and coordinates are used.",
        "KMZ": "Zipped KML file. Same as KML after unzipping.",
    }
    for tab, (fmt, name) in zip(tabs, SAMPLE_FILES):
        with tab:
            st.caption(notes[fmt])
            if (DATA_DIR / name).exists():
                try:
                    st.dataframe(load_sample_raw(name).head(50), use_container_width=True, hide_index=True)
                    st.caption("Showing the first 50 records exactly as they appear in the source file.")
                except Exception as exc:
                    st.error(f"Could not read {name}: {exc}")
            else:
                st.warning(f"`data/{name}` not found.")

    st.subheader("Common schema (15 fields)")
    st.dataframe(pd.DataFrame([
        {"Field": f, "Type": FIELD_DESCRIPTIONS[f][0], "Description": FIELD_DESCRIPTIONS[f][1]}
        for f in COMMON_COLUMNS
    ]), use_container_width=True, hide_index=True)

    st.subheader("Source field mapping")
    st.caption("Which source column feeds each common field. “—” means the source does not provide that field.")
    mapping_table = pd.DataFrame({
        fmt: [SCHEMA_MAPPING[fmt].get(f) or "—" for f in COMMON_COLUMNS]
        for fmt in SCHEMA_MAPPING
    }, index=COMMON_COLUMNS)
    mapping_table.index.name = "Common field"
    st.dataframe(mapping_table, use_container_width=True)

    st.info("If a header does not match exactly, the app also tries normalized names and common aliases "
            "(e.g. “postal code”, “zipcode”). Fields that do not exist in a source are left missing, never invented.")


# ---------------------------------------------------------------------
# 4. PREPROCESSING
# ---------------------------------------------------------------------
def page_preprocessing():
    st.header("🧹 Preprocessing")
    tab_rules, tab_results = st.tabs(["📖 Rules & method", "📊 Results (after implementation)"])

    with tab_rules:
        st.write("Cleaning follows the rules of the source notebook. Descriptive fields such as phone numbers and "
                 "addresses are never fabricated. Preprocessing starts from the EDA output (missing-value tokens already standardized).")
        st.dataframe(pd.DataFrame([
            {"#": 1, "Step": "Trim text", "Rule": "Remove leading/trailing spaces from text fields"},
            {"#": 2, "Step": "County status", "Rule": "“in county” → In-County; “out county” → Out-of-County"},
            {"#": 3, "Step": "Hospital type", "Rule": "general → General Hospital; emergency → Emergency Center; specialty → Specialty Hospital"},
            {"#": 4, "Step": "Numeric conversion", "Rule": "Longitude, Latitude, Data_Year, Bed_Count, Quality_Score converted to numbers; bad text → missing"},
            {"#": 5, "Step": "Invalid beds", "Rule": "Negative Bed_Count → missing"},
            {"#": 6, "Step": "Invalid quality score", "Rule": "Quality_Score outside 0–100 → missing"},
            {"#": 7, "Step": "Coordinate check", "Rule": "Latitude outside −90..90 or Longitude outside −180..180 → missing"},
            {"#": 8, "Step": "ZIP validation", "Rule": "Keep only 5-digit ZIP codes (removes “.0” suffix first)"},
            {"#": 9, "Step": "Phone standardization", "Rule": "Keep digits only; valid only if exactly 10 digits"},
            {"#": 10, "Step": "Name correction", "Rule": "“National Institue” → “National Institute”"},
            {"#": 11, "Step": "Duplicate removal", "Rule": "Remove exact duplicate rows (Source_Record_ID is ignored when comparing)"},
            {"#": 12, "Step": "Median imputation", "Rule": "Fill missing Bed_Count and Quality_Score with the column median"},
        ]), use_container_width=True, hide_index=True)

        st.subheader("Validation checks after cleaning")
        st.dataframe(pd.DataFrame([
            {"Check": "Negative Bed Count", "Passes when": "No negative bed counts"},
            {"Check": "Quality Score 0-100", "Passes when": "All scores within 0–100"},
            {"Check": "Latitude Range", "Passes when": "All latitudes within −90..90"},
            {"Check": "Longitude Range", "Passes when": "All longitudes within −180..180"},
            {"Check": "Duplicate Rows", "Passes when": "No duplicate rows remain"},
            {"Check": "Invalid ZIP format", "Passes when": "All non-missing ZIPs are 5 digits"},
            {"Check": "Invalid phone format", "Passes when": "All non-missing phones are 10 digits"},
        ]), use_container_width=True, hide_index=True)

    with tab_results:
        result = require_result()
        if not result:
            return
        clean = result["clean"]
        st.subheader("Before vs after")
        st.dataframe(result["before_after"], use_container_width=True, hide_index=True)
        st.subheader("Cleaned data")
        st.dataframe(clean.fillna("NA"), use_container_width=True, hide_index=True, height=430)
        st.subheader("Preprocessing log")
        st.dataframe(result["actions"], use_container_width=True, hide_index=True)
        st.subheader("Median imputation report")
        st.dataframe(result["imputation"], use_container_width=True, hide_index=True)
        st.subheader("Final record count by source")
        st.dataframe(result["final_source_summary"], use_container_width=True, hide_index=True)
        st.download_button("Download cleaned CSV", clean.fillna("NA").to_csv(index=False).encode("utf-8-sig"),
                           "hospital_final_clean_data.csv", "text/csv")


# ---------------------------------------------------------------------
# 5. EDA
# ---------------------------------------------------------------------
def page_eda():
    st.header("🔎 Exploratory Data Analysis (EDA)")
    tab_method, tab_results = st.tabs(["📖 EDA steps", "📊 Results (after implementation)"])

    with tab_method:
        st.write("EDA is performed on the **combined, schema-standardized data before cleaning**, "
                 "so that real data-quality problems are visible. Tokens such as `NA`, `N/A`, `NULL`, `None` and "
                 "`unknown` are first converted to real missing values.")
        st.dataframe(pd.DataFrame([
            {"EDA step": "Dataset shape", "What it does": "Number of rows and columns of the combined data", "Output": "Metrics"},
            {"EDA step": "Data-type review", "What it does": "Shows the data type of every column", "Output": "Table"},
            {"EDA step": "Missing-value analysis", "What it does": "Missing count and % per column, total per source, and a source × column heat map", "Output": "Tables + bar charts + heat map"},
            {"EDA step": "Duplicate analysis", "What it does": "Exact duplicate rows and potential duplicates (same name + address + city)", "Output": "Counts + record tables"},
            {"EDA step": "Unique-value analysis", "What it does": "Number of distinct values per column", "Output": "Table"},
            {"EDA step": "Category analysis", "What it does": "Value counts for County_Status, Hospital_Type, Record_Status and Source_Format", "Output": "Tables + bar charts"},
            {"EDA step": "Descriptive statistics", "What it does": "Count, mean, std, min, quartiles, max for numeric fields", "Output": "Summary table"},
            {"EDA step": "Invalid-value detection", "What it does": "Negative beds, quality score outside 0-100, impossible coordinates", "Output": "Counts + record tables"},
            {"EDA step": "Outlier detection (IQR)", "What it does": "Values outside Q1 − 1.5·IQR and Q3 + 1.5·IQR for beds, quality, latitude, longitude", "Output": "Table + box plots"},
            {"EDA step": "Distributions", "What it does": "Histograms of Bed_Count and Quality_Score, source and county distributions", "Output": "Charts"},
            {"EDA step": "EDA issue log", "What it does": "One table listing every problem found and the action required", "Output": "Table (also exported to Excel)"},
        ]), use_container_width=True, hide_index=True)
        st.subheader("About the IQR outlier rule")
        st.markdown("A value is treated as an outlier if it is below **Q1 − 1.5 × IQR** or above **Q3 + 1.5 × IQR**, "
                    "where IQR = Q3 − Q1. Outliers are reported for review; they are not deleted.")

    with tab_results:
        result = require_result()
        if not result:
            return
        eda = result["eda"]
        frame = result["raw"]

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Rows", f"{eda['rows']:,}")
        m2.metric("Columns", eda["columns"])
        m3.metric("Missing cells", f"{int(eda['missing']['Missing_Count'].sum()):,}")
        m4.metric("Exact duplicates", eda["exact_dups"])
        m5.metric("Potential duplicates", eda["potential_dups"])

        t_miss, t_dup, t_cat, t_num, t_out, t_dist, t_issue = st.tabs([
            "Missing values", "Duplicates & uniqueness", "Categories", "Numeric & invalid values",
            "Outliers", "Distributions", "Issue log"])

        with t_miss:
            c1, c2 = st.columns(2)
            with c1:
                miss = eda["missing"].query("Missing_Count > 0").sort_values("Missing_Count")
                if miss.empty:
                    st.success("No missing values were detected in the standardized data.")
                else:
                    st.plotly_chart(px.bar(miss, x="Missing_Count", y="Column", orientation="h",
                                           hover_data=["Missing_Percentage"],
                                           title="Missing values by column"), use_container_width=True)
            with c2:
                st.plotly_chart(px.bar(eda["missing_by_source"], x="Source_Format", y="Total_Missing_Cells",
                                       color="Source_Format", text_auto=True,
                                       title="Missing cells by source format"), use_container_width=True)
            st.markdown("**Missing % by source and column**")
            st.plotly_chart(px.imshow(eda["missing_heatmap"], text_auto=".0f", aspect="auto",
                                      color_continuous_scale="Blues",
                                      labels=dict(color="% missing")), use_container_width=True)
            st.caption("Fields a source does not provide (for example Address in KML) show as 100 %. That is expected, not an error.")
            a, b = st.columns(2)
            with a:
                st.markdown("**Missing-value report**")
                st.dataframe(eda["missing"], use_container_width=True, hide_index=True)
            with b:
                st.markdown("**Data types**")
                st.dataframe(eda["dtypes"], use_container_width=True, hide_index=True)

        with t_dup:
            d1, d2 = st.columns(2)
            d1.metric("Exact duplicate rows", eda["exact_dups"])
            d2.metric("Potential hospital duplicates (name + address + city)", eda["potential_dups"])
            if IGNORE_ID_COLUMNS_FOR_DUPLICATES:
                st.caption("Exact duplicates are compared on all hospital fields plus Source_Format; the per-file "
                           "Source_Record_ID is ignored, otherwise no two rows could ever match.")
            with st.expander(f"Exact duplicate records ({len(eda['dup_rows'])} rows including the first occurrence)"):
                st.dataframe(eda["dup_rows"], use_container_width=True, hide_index=True)
            with st.expander(f"Potential duplicate records ({len(eda['potential_rows'])} rows)"):
                st.dataframe(eda["potential_rows"], use_container_width=True, hide_index=True)
            st.markdown("**Unique values per column**")
            st.dataframe(eda["unique"], use_container_width=True, hide_index=True)

        with t_cat:
            st.caption("Counts are before cleaning, so spelling and case variants (for example “in county” and “In-County”) "
                       "show up as separate categories.")
            for col in CATEGORICAL_COLUMNS:
                counts = eda["categorical"][col]
                st.markdown(f"**{col}**")
                left, right = st.columns([1, 1.6])
                with left:
                    st.dataframe(counts, use_container_width=True, hide_index=True)
                with right:
                    st.plotly_chart(px.bar(counts, x=col, y="Count", text_auto=True), use_container_width=True)

        with t_num:
            st.markdown("**Descriptive statistics (numeric fields)**")
            st.dataframe(eda["numeric_summary"], use_container_width=True)
            i1, i2, i3 = st.columns(3)
            i1.metric("Negative bed-count records", len(eda["invalid_beds"]))
            i2.metric("Invalid quality-score records", len(eda["invalid_quality"]))
            i3.metric("Invalid coordinate records", len(eda["invalid_coords"]))
            for label, key in [("Negative bed-count records", "invalid_beds"),
                               ("Invalid quality-score records (outside 0-100)", "invalid_quality"),
                               ("Invalid coordinate records", "invalid_coords")]:
                with st.expander(label):
                    if eda[key].empty:
                        st.success("None found.")
                    else:
                        st.dataframe(eda[key], use_container_width=True, hide_index=True)

        with t_out:
            st.markdown("**IQR outlier analysis**")
            st.dataframe(eda["outliers"], use_container_width=True, hide_index=True)
            plot_cols = st.columns(2)
            for i, col in enumerate(OUTLIER_COLUMNS):
                with plot_cols[i % 2]:
                    data = frame[col].dropna()
                    if data.empty:
                        st.info(f"No valid data for {col}, so no box plot is shown.")
                    else:
                        st.plotly_chart(px.box(frame, y=col, points="outliers", hover_data=["Hospital_Name", "Source_Format"],
                                               title=f"Outlier analysis - {col}"), use_container_width=True)
            for col in OUTLIER_COLUMNS:
                rows = eda["outlier_records"][col]
                with st.expander(f"{col}: {len(rows)} outlier record(s)"):
                    if rows.empty:
                        st.success("No IQR outliers.")
                    else:
                        st.dataframe(rows.head(50), use_container_width=True, hide_index=True)

        with t_dist:
            h1, h2 = st.columns(2)
            with h1:
                if frame["Bed_Count"].notna().any():
                    st.plotly_chart(px.histogram(frame, x="Bed_Count", nbins=30,
                                                 title="Distribution of hospital bed count"), use_container_width=True)
                else:
                    st.info("No Bed_Count values to plot.")
            with h2:
                if frame["Quality_Score"].notna().any():
                    st.plotly_chart(px.histogram(frame, x="Quality_Score", nbins=30,
                                                 title="Distribution of quality score"), use_container_width=True)
                else:
                    st.info("No Quality_Score values to plot.")
            s1, s2 = st.columns(2)
            with s1:
                st.plotly_chart(px.pie(eda["categorical"]["Source_Format"], names="Source_Format", values="Count",
                                       hole=.45, title="Records by source format"), use_container_width=True)
            with s2:
                st.plotly_chart(px.bar(eda["categorical"]["County_Status"], x="County_Status", y="Count",
                                       text_auto=True, title="County status distribution"), use_container_width=True)

        with t_issue:
            st.markdown("**EDA issue log**")
            if eda["issue_log"].empty:
                st.success("No data-quality issues were found.")
            else:
                st.dataframe(eda["issue_log"], use_container_width=True, hide_index=True)
            st.caption("Each issue is handled in the Preprocessing step, except potential duplicates and outliers, which are only reported.")


# ---------------------------------------------------------------------
# 6. IMPLEMENTATION  (formerly "Execution")
# ---------------------------------------------------------------------
def page_implementation():
    st.header("▶️ Implementation")
    st.write("Run the full pipeline: parse → map to common schema → combine → EDA → preprocess → validate → export. "
             "Files are processed in memory and are not saved to disk.")

    source = st.radio("Data source", ["Use the project dataset files", "Upload my own files"], horizontal=True)

    files = []
    if source == "Use the project dataset files":
        files = sample_files()
        if files:
            st.success(f"{len(files)} dataset file(s) ready: " + ", ".join(f.name for f in files))
        else:
            st.error("No dataset files found. Make sure the `data/` folder with the six files is in your repository.")
    else:
        files = st.file_uploader(
            "Choose one or more files",
            type=["csv", "json", "xml", "html", "htm", "kml", "kmz"],
            accept_multiple_files=True,
            help="Supported formats: CSV, JSON, XML, HTML, KML, KMZ.",
        ) or []

    if st.button("▶️ Run pipeline", type="primary", disabled=not files):
        if process_files(files):
            st.success("Pipeline completed. Results are shown below, and in the **EDA** and **Preprocessing** pages.")

    result = st.session_state.get("hospital_result")
    if not result or "eda" not in result:
        return

    st.divider()
    t_over, t_eda, t_data, t_val, t_exp = st.tabs(
        ["📊 Overview", "🔎 EDA findings", "🗂️ Integrated data", "✅ Validation", "⬇️ Export"])

    with t_over:
        raw = result["raw"]
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Files processed", len(result["parsed"]))
        k2.metric("Combined records", f"{len(raw):,}")
        k3.metric("Schema fields", len(COMMON_COLUMNS))
        k4.metric("Missing cells before", f"{result['before_after'].loc[result['before_after']['Metric']=='Missing cells','Before_Preprocessing'].iloc[0]:,}")
        k5.metric("Validation checks passed", f"{int((result['validation']['Status']=='PASS').sum())}/{len(result['validation'])}")
        left, right = st.columns([1.05, 1])
        with left:
            st.subheader("Records by file")
            st.dataframe(result["source_summary"], use_container_width=True, hide_index=True)
            fig = px.bar(result["source_summary"], x="File", y="Input_Rows", color="Source_Format",
                         title="Input records per file", text_auto=True)
            st.plotly_chart(fig, use_container_width=True)
        with right:
            st.subheader("Before vs after")
            st.dataframe(result["before_after"], use_container_width=True, hide_index=True)
            st.subheader("Final records by source")
            st.dataframe(result["final_source_summary"], use_container_width=True, hide_index=True)
        st.caption("Counts are calculated from the files processed in this session.")

    with t_eda:
        eda = result["eda"]
        e1, e2, e3, e4, e5 = st.columns(5)
        e1.metric("Rows analysed", f"{eda['rows']:,}")
        e2.metric("Missing cells", f"{int(eda['missing']['Missing_Count'].sum()):,}")
        e3.metric("Exact duplicates", eda["exact_dups"])
        e4.metric("Invalid values", len(eda["invalid_beds"]) + len(eda["invalid_quality"]) + len(eda["invalid_coords"]))
        e5.metric("IQR outliers", int(eda["outliers"]["Outlier_Count"].sum()))
        st.subheader("EDA issue log")
        if eda["issue_log"].empty:
            st.success("No data-quality issues were found.")
        else:
            st.dataframe(eda["issue_log"], use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        with c1:
            miss = eda["missing"].query("Missing_Count > 0").sort_values("Missing_Count")
            if not miss.empty:
                st.plotly_chart(px.bar(miss, x="Missing_Count", y="Column", orientation="h",
                                       title="Missing values by column"), use_container_width=True)
        with c2:
            st.subheader("IQR outlier summary")
            st.dataframe(eda["outliers"], use_container_width=True, hide_index=True)
        st.info("Open the **🔎 EDA** page in the sidebar for the full analysis: duplicates, categories, "
                "descriptive statistics, box plots and histograms.")

    with t_data:
        raw = result["raw"]
        st.caption("Combined data after schema mapping and missing-token standardization, before preprocessing.")
        search = st.text_input("Filter rows by text", key="raw_filter")
        view = raw
        if search:
            mask = raw.astype("string").apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
            view = raw.loc[mask]
        st.dataframe(view, use_container_width=True, hide_index=True, height=460)
        st.download_button("Download standardized CSV", raw.to_csv(index=False).encode("utf-8-sig"),
                           "hospital_standardized_raw.csv", "text/csv")

    with t_val:
        st.dataframe(result["validation"], use_container_width=True, hide_index=True)
        failures = result["validation"].query("Status == 'FAIL'")
        if failures.empty:
            st.success("All configured validation checks passed for this dataset.")
        else:
            st.warning(f"{len(failures)} validation check(s) still report invalid values. Review them before using the dataset.")
        st.caption("A PASS only means the specific check passed. Missing descriptive fields can remain in the final data by design.")

    with t_exp:
        st.write("Download a formatted Excel workbook with the standardized data, cleaned data, schema mapping, "
                 "missing-value report, EDA issue log, outlier analysis, before/after comparison, preprocessing log, "
                 "validation report, imputation report and final source summary.")
        excel_bytes = create_excel_report(
            result["raw"], result["clean"], result["source_summary"], result["mapping"], result["missing"],
            result["validation"], result["imputation"], result["actions"], result["before_after"],
            result["dtypes"], result["outliers"], result["eda"]["issue_log"], result["final_source_summary"],
        )
        st.download_button(
            "📥 Download full Excel report", data=excel_bytes,
            file_name="Hospital_Data_Final_EDA_Preprocessing.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.download_button("Download schema mapping CSV", result["mapping"].to_csv(index=False).encode("utf-8-sig"),
                           "hospital_schema_mapping.csv", "text/csv")
        st.download_button("Download validation CSV", result["validation"].to_csv(index=False).encode("utf-8-sig"),
                           "hospital_validation_report.csv", "text/csv")
        st.download_button("Download EDA issue log CSV", result["eda"]["issue_log"].to_csv(index=False).encode("utf-8-sig"),
                           "hospital_eda_issue_log.csv", "text/csv")


# ---------------------------------------------------------------------
# Sidebar navigation  (Architecture sits above About Dataset)
# ---------------------------------------------------------------------
PAGES = {
    "ℹ️  About Project": page_about_project,
    "🏗️  Architecture": page_architecture,
    "📚  About Dataset": page_about_dataset,
    "🧹  Preprocessing": page_preprocessing,
    "🔎  EDA": page_eda,
    "▶️  Implementation": page_implementation,
}

with st.sidebar:
    st.markdown("""
<div class="side-brand">
  <div class="logo">🏥</div>
  <div>
    <div class="title">Hospital Data<br>Integration Studio</div>
    <div class="sub">INTEGRATE · CLEAN · EXPORT</div>
  </div>
</div>
<div class="side-label">Navigation</div>
""", unsafe_allow_html=True)

    choice = st.radio("Navigate", list(PAGES.keys()), label_visibility="collapsed")

    st.divider()
    if st.session_state.get("hospital_result"):
        n_files = len(st.session_state["hospital_result"]["parsed"])
        st.markdown(f'<div class="status-pill"><span class="dot on"></span>'
                    f'<span>Pipeline ready · {n_files} file(s) processed</span></div>',
                    unsafe_allow_html=True)
    else:
        st.markdown('<div class="status-pill"><span class="dot off"></span>'
                    '<span>Pipeline not run yet</span></div>', unsafe_allow_html=True)
    st.markdown('<div class="side-foot">Tip: run <b>Implementation</b> first to see live results in the '
                'EDA and Preprocessing pages.</div>', unsafe_allow_html=True)

PAGES[choice]()
