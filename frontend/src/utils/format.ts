/**
 * Scale-factor-aware currency formatter.
 *
 * Dashboard values are pre-divided by scaleFactor before being sent to the
 * frontend, so the raw dollar amount is: value * scaleFactor.
 * e.g. scaleFactor=1000, value=23291.93 → actual = $23,291,930 → "$23.3M"
 */
export function formatCurrency(value: number, scaleFactor: number): string {
  const actual = value * scaleFactor;
  const isNegative = actual < 0;
  const abs = Math.abs(actual);
  let formatted: string;
  if (abs >= 1_000_000_000) formatted = `$${(abs / 1_000_000_000).toFixed(1)}B`;
  else if (abs >= 1_000_000) formatted = `$${(abs / 1_000_000).toFixed(1)}M`;
  else if (abs >= 1_000) formatted = `$${(abs / 1_000).toFixed(1)}K`;
  else formatted = `$${abs.toFixed(0)}`;
  return isNegative ? `(${formatted})` : formatted;
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
  const isNegative = value < 0;
  const formatted = `${(Math.abs(value) * 100).toFixed(1)}%`;
  return isNegative ? `(${formatted})` : formatted;
}
