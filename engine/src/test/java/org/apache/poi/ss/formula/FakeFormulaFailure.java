package org.apache.poi.ss.formula;

/** Test double for POI FormulaFailure to exercise formula exception handling. */
public final class FakeFormulaFailure extends RuntimeException {
  private static final long serialVersionUID = 1L;

  public FakeFormulaFailure(String message) {
    super(message);
  }
}
