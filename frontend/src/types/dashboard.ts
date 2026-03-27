export interface WaterfallData {
  period_label: string;
  bop: number;
  new_logo: number;
  upsell: number;
  downsell: number;
  churn: number;
  eop: number;
}

export interface StatsData {
  total_arr: number;
  customer_count: number;
  net_retention_pct: number | null;
  yoy_growth_pct: number | null;
  lost_only_retention_pct: number | null;
  punitive_retention_pct: number | null;
}

export interface TopCustomer {
  rank: number;
  name: string;
  arr: number;
  change_pct: number | null;
  pct_of_total: number;
  trend: number[];
  status: string;
  attributes: Record<string, string>;
  cohort: string;
}

export interface OverviewData {
  periods: string[];
  arr_over_time: number[];
  arr_growth_pcts: (number | null)[];
  waterfall: WaterfallData | null;
  stats: StatsData;
  top_customers: TopCustomer[];
  latest_period_label: string;
  latest_period_date: string;
}

export interface CohortEntry {
  label: string;
  count: number;
  starting_arr: number;
  arr: (number | null)[];
  customers: (number | null)[];
  ndr: (number | null)[];
  logo_retention: (number | null)[];
}

export interface CohortData {
  periods: string[];
  cohorts: CohortEntry[];
}

export interface AttributeOption {
  name: string;
  values: string[];
}

export interface DashboardResponse {
  overview: OverviewData;
  cohort: CohortData;
  granularity: string;
  available_granularities: string[];
  scale_factor: number;
  attribute_options: AttributeOption[];
}

export type CohortMetric = "ndr" | "arr" | "logo_retention" | "customers";
