/**
 * filter.js
 * ---------
 * Core filtering logic — JavaScript port of filter_ids.py.
 *
 * Column positions (1-based, matching the Python original)
 *   Col 1  – ID
 *   Col 3  – Status  → must be "DONE" or "CLOSED"
 *   Col 5  – Group   → must be "HIS_GROUP"
 *   Col 6  – Type    → must be "Defect (Bug)"
 */

const COL_ID     = 1;
const COL_STATUS = 3;
const COL_GROUP  = 5;
const COL_TYPE   = 6;

const VALID_STATUSES = new Set(["DONE", "CLOSED"]);
const VALID_GROUP    = "HIS_GROUP";
const VALID_TYPE     = "Defect (Bug)";

/**
 * Returns true when a row (0-based array) satisfies all filter conditions.
 * @param {any[]} row
 * @returns {boolean}
 */
export function matches(row) {
  if (row.length < COL_TYPE) return false;

  const status = row[COL_STATUS - 1] != null ? String(row[COL_STATUS - 1]).trim() : "";
  const group  = row[COL_GROUP  - 1] != null ? String(row[COL_GROUP  - 1]).trim() : "";
  const type   = row[COL_TYPE   - 1] != null ? String(row[COL_TYPE   - 1]).trim() : "";

  return VALID_STATUSES.has(status) && group === VALID_GROUP && type === VALID_TYPE;
}

/**
 * Extracts IDs from an array of rows (each row is a plain array of cell values).
 * @param {any[][]} rows
 * @param {boolean} skipHeader  – when true, the first row is treated as a header and skipped
 * @returns {any[]}
 */
export function extractIds(rows, skipHeader) {
  const dataRows = skipHeader && rows.length > 0 ? rows.slice(1) : rows;
  const result = [];

  for (const row of dataRows) {
    if (matches(row)) {
      const id = row[COL_ID - 1];
      if (id != null) result.push(id);
    }
  }

  return result;
}
