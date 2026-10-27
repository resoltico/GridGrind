package dev.erst.gridgrind.core.model.formula;

/** High-level capability classification for one authored or loaded formula cell. */
public enum FormulaCapabilityKind {
  EVALUABLE_NOW,
  UNEVALUABLE_NOW,
  UNPARSEABLE_BY_POI
}
