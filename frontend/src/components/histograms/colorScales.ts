/** Shared color scale for retention metric cells (red → green around 100%). */
export function retentionColor(v: number): { bg: string; text: string } {
  if (v >= 1.5)  return { bg: "rgba(5, 150, 105, 0.38)", text: "#047857" };
  if (v >= 1.2)  return { bg: "rgba(5, 150, 105, 0.28)", text: "#059669" };
  if (v >= 1.1)  return { bg: "rgba(16, 185, 129, 0.22)", text: "#059669" };
  if (v >= 1.0)  return { bg: "rgba(16, 185, 129, 0.14)", text: "#10B981" };
  if (v >= 0.9)  return { bg: "rgba(245, 158, 11, 0.14)", text: "#D97706" };
  if (v >= 0.8)  return { bg: "rgba(249, 115, 22, 0.16)", text: "#EA580C" };
  return { bg: "rgba(239, 68, 68, 0.20)", text: "#DC2626" };
}
