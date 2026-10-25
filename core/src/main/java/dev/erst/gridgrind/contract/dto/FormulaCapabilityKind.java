package dev.erst.gridgrind.contract.dto;

/** High-level capability classification for one authored or loaded formula cell. */
public enum FormulaCapabilityKind {
  EVALUABLE_NOW,
  UNEVALUABLE_NOW,
  UNPARSEABLE_BY_POI
}
