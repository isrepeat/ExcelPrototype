export function eventCellValue(value) {
  if (value == null || value === 0) return null;
  if (typeof value === "string" && (value.trim() === "" || /^[+-]?0+(?:[.,]0+)?$/.test(value.trim()))) return null;
  return value;
}