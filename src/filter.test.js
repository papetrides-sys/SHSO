/**
 * filter.test.js
 * --------------
 * Unit tests for filter.js — mirrors the Python test_filter_ids.py suite.
 */

import { describe, test, expect } from "@jest/globals";
import { matches, extractIds } from "./filter.js";

// ---------------------------------------------------------------------------
// matches()
// ---------------------------------------------------------------------------
describe("matches()", () => {
  const goodRow = ["ID-01", "irrelevant", "DONE", "irrelevant", "HIS_GROUP", "Defect (Bug)"];

  test("returns true for a fully matching row", () => {
    expect(matches(goodRow)).toBe(true);
  });

  test("accepts CLOSED as a valid status", () => {
    const row = [...goodRow];
    row[2] = "CLOSED";
    expect(matches(row)).toBe(true);
  });

  test("returns false when status is wrong", () => {
    const row = [...goodRow];
    row[2] = "OPEN";
    expect(matches(row)).toBe(false);
  });

  test("returns false when group is wrong", () => {
    const row = [...goodRow];
    row[4] = "OTHER_GROUP";
    expect(matches(row)).toBe(false);
  });

  test("returns false when type is wrong", () => {
    const row = [...goodRow];
    row[5] = "Enhancement";
    expect(matches(row)).toBe(false);
  });

  test("returns false when row is too short", () => {
    expect(matches(["ID-01", "a", "DONE"])).toBe(false);
  });

  test("trims whitespace from cell values", () => {
    const row = ["ID-01", "x", "  DONE  ", "x", "  HIS_GROUP  ", "  Defect (Bug)  "];
    expect(matches(row)).toBe(true);
  });

  test("treats null cells as empty strings (no match)", () => {
    const row = [null, null, null, null, null, null];
    expect(matches(row)).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// extractIds()
// ---------------------------------------------------------------------------
describe("extractIds()", () => {
  const makeRow = (id, status, group, type) => [
    id, "x", status, "x", group, type,
  ];

  const rows = [
    makeRow("A1", "DONE",   "HIS_GROUP", "Defect (Bug)"),
    makeRow("A2", "CLOSED", "HIS_GROUP", "Defect (Bug)"),
    makeRow("A3", "OPEN",   "HIS_GROUP", "Defect (Bug)"),   // status mismatch
    makeRow("A4", "DONE",   "OTHER",     "Defect (Bug)"),   // group mismatch
    makeRow("A5", "DONE",   "HIS_GROUP", "Enhancement"),    // type mismatch
  ];

  test("returns IDs of rows that pass all conditions", () => {
    expect(extractIds(rows, false)).toEqual(["A1", "A2"]);
  });

  test("skips header row when skipHeader=true", () => {
    const withHeader = [
      ["ID", "Col2", "Status", "Col4", "Group", "Type"], // header
      ...rows,
    ];
    expect(extractIds(withHeader, true)).toEqual(["A1", "A2"]);
  });

  test("returns empty array when nothing matches", () => {
    const noMatch = [makeRow("X1", "OPEN", "HIS_GROUP", "Defect (Bug)")];
    expect(extractIds(noMatch, false)).toEqual([]);
  });

  test("returns empty array for empty input", () => {
    expect(extractIds([], false)).toEqual([]);
  });

  test("skipping header on single-row input returns empty array", () => {
    const headerOnly = [["ID", "x", "Status", "x", "Group", "Type"]];
    expect(extractIds(headerOnly, true)).toEqual([]);
  });

  test("excludes rows where ID is null", () => {
    const row = [null, "x", "DONE", "x", "HIS_GROUP", "Defect (Bug)"];
    expect(extractIds([row], false)).toEqual([]);
  });
});
