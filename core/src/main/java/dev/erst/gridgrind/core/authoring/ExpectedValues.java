package dev.erst.gridgrind.core.authoring;

import dev.erst.gridgrind.core.model.value.CellValue;

/** Authored expected scalar values using the canonical resolved value model. */
public final class ExpectedValues {
  private ExpectedValues() {}

  /** Returns an expected blank value. */
  public static CellValue blank() {
    return new CellValue.Blank();
  }

  /** Returns an expected exact text value. */
  public static CellValue text(String value) {
    return new CellValue.Text(value);
  }

  /** Returns an expected numeric value; the assertion enforces authored finiteness. */
  public static CellValue number(double value) {
    return new CellValue.NumberValue(value);
  }

  /** Returns an expected exact boolean value. */
  public static CellValue bool(boolean value) {
    return new CellValue.BooleanValue(value);
  }

  /** Returns an expected exact reported error literal. */
  public static CellValue error(String value) {
    return new CellValue.ErrorValue(value);
  }
}
