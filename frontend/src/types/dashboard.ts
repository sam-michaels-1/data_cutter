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
  annualized_lost_only_retention_pct: number | null;
  annualized_punitive_retention_pct: number | null;
  annualized_net_retention_pct: number | null;
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
  multiSelect?: boolean;
}

export type FilterValue = string | string[];
export type Filters = Record<string, FilterValue>;

export interface DashboardResponse {
  overview: OverviewData;
  cohort: CohortData;
  granularity: string;
  available_granularities: string[];
  scale_factor: number;
  attribute_options: AttributeOption[];
  data_type: string;
}

export type CohortMetric = "ndr" | "arr" | "logo_retention" | "customers";

export type CustomerEventType =
  | "new"
  | "upsell"
  | "downsell"
  | "churn"
  | "reactivation"
  | "flat"
  | "inactive";

export interface CustomerPeriodEntry {
  period_label: string;
  arr: number;
  change: number | null;
  change_pct: number | null;
  event: CustomerEventType;
}

export interface CustomerMovement {
  period_label: string;
  type: Exclude<CustomerEventType, "flat" | "inactive">;
  amount: number;
}

export interface AttributeTransition {
  attribute: string;
  from: string;
  to: string;
  date: string;
  period_label: string;
}

export interface CustomerListEntry {
  name: string;
  current_arr: number;
}

export interface CustomerDetail {
  name: string;
  attributes: Record<string, string>;
  cohort: string;
  status: string;
  current_arr: number;
  first_arr: number;
  first_period_label: string;
  peak_arr: number;
  peak_period_label: string;
  lifetime_total: number;
  pct_of_total: number;
  total_change: number | null;
  total_change_pct: number | null;
  yoy_change_pct: number | null;
  cagr: number | null;
  timeline: CustomerPeriodEntry[];
  movements: CustomerMovement[];
  attribute_transitions: AttributeTransition[];
  granularity: string;
  available_granularities: string[];
  scale_factor: number;
  data_type: string;
}
