"""
Pydantic request / response models for the wizard API.
"""
from typing import Optional

from pydantic import BaseModel


# ---------- Upload ----------

class UploadResponse(BaseModel):
    session_id: str
    filename: str
    sheet_names: list[str]


# ---------- Detect Columns ----------

class DetectColumnsRequest(BaseModel):
    session_id: str
    sheet_name: str


class ColumnInfo(BaseModel):
    letter: str
    header: str
    sample_values: list[str]


class AttributeCol(BaseModel):
    header: str
    letter: str


class DetectedMapping(BaseModel):
    date_col: str
    customer_id_col: str
    arr_col: str
    attribute_cols: list[AttributeCol]


class DetectColumnsResponse(BaseModel):
    columns: list[ColumnInfo]
    detected_mapping: DetectedMapping
    row_count: int
    auto_scale_factor: int
    detected_frequency: Optional[str] = None  # "monthly", "quarterly", or "annual"


# ---------- Generate ----------

class ColumnMapping(BaseModel):
    date_col: str
    customer_id_col: str
    arr_col: str


class AttributeSelection(BaseModel):
    display_name: str
    letter: str


class GenerateRequest(BaseModel):
    session_id: str
    sheet_name: str
    data_type: str  # "arr" | "revenue"
    data_frequency: Optional[str] = None  # "monthly" | "quarterly" — user override
    column_mapping: ColumnMapping
    attributes: list[AttributeSelection]
    output_granularities: list[str]   # e.g. ["monthly", "quarterly", "annual"]
    fiscal_year_end_month: int        # 1-12


class GenerateResponse(BaseModel):
    status: str
    download_id: str
