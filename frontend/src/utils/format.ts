/**
 * Scale-factor-aware currency formatter.
 *
 * Dashboard values are pre-divided by scaleFactor before being sent to the
 * frontend, so the raw dollar amount is: value * scaleFactor.
 * e.g. scaleFactor=1000, value=23291.93 → actual = $23,291,930 → "$23.3M"
 */
export function formatCurrency(value: number, scaleFactor: number): string {
  const actual = value * scaleFactor;
  if (actual >= 1_000_000_000) return `$${(actual / 1_000_000_000).toFixed(1)}B`;
  if (actual >= 1_000_000) return `$${(actual / 1_000_000).toFixed(1)}M`;
  if (actual >= 1_000) return `$${(actual / 1_000).toFixed(1)}K`;
  return `$${actual.toFixed(0)}`;
}

/**
 * Y-axis tick formatter — same logic, rounds to 1 decimal.
 */
export function formatYAxis(value: number, scaleFactor: number): string {
  return formatCurrency(value, scaleFactor);
}

/**
 * Format a percentage (0–1 range) to 1 decimal place.
 */
export function formatPct(value: number | null): string {
  if (value == null) return "N/A";
  return `${(value * 100).toFixed(1)}%`;
}
