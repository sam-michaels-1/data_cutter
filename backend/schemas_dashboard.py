"""Pydantic models for the dashboard API."""
from __future__ import annotations

from typing import Dict, List, Optional

from pydantic import BaseModel


class WaterfallData(BaseModel):
    period_label: str
    bop: float
    new_logo: float
    upsell: float
    downsell: float
    churn: float
    eop: float


class StatsData(BaseModel):
    total_arr: float
    customer_count: int
    net_retention_pct: Optional[float]
    yoy_growth_pct: Optional[float]
    lost_only_retention_pct: Optional[float]
    punitive_retention_pct: Optional[float]


class TopCustomer(BaseModel):
    rank: int
    name: str
    arr: float
    change_pct: Optional[float]
    pct_of_total: float
    trend: List[float]
    status: str
    attributes: Dict[str, str]
    cohort: str


class OverviewData(BaseModel):
    periods: List[str]
    arr_over_time: List[float]
    arr_growth_pcts: List[Optional[float]]
    waterfall: Optional[WaterfallData]
    stats: StatsData
    top_customers: List[TopCustomer]
    latest_period_label: str
    latest_period_date: str


class CohortEntry(BaseModel):
    label: str
    count: int
    starting_arr: float
    arr: List[Optional[float]]
    customers: List[Optional[int]]
    ndr: List[Optional[float]]
    logo_retention: List[Optional[float]]


class CohortData(BaseModel):
    periods: List[str]
    cohorts: List[CohortEntry]


class AttributeOption(BaseModel):
    name: str
    values: List[str]


class DashboardResponse(BaseModel):
    overview: OverviewData
    cohort: CohortData
    granularity: str
    available_granularities: List[str]
    scale_factor: int
    attribute_options: List[AttributeOption]
    data_type: str = "arr"
