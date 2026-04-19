package dev.erst.gridgrind.excel;

/** Engine-side capability classification for one formula cell under the current runtime. */
public enum ExcelFormulaCapabilityKind {
  EVALUABLE_NOW,
  UNEVALUABLE_NOW,
  UNPARSEABLE_BY_POI
}
