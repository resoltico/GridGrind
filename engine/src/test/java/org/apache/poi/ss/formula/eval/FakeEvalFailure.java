package org.apache.poi.ss.formula.eval;

/** Test double for POI EvalFailure to exercise formula evaluation exception handling. */
public final class FakeEvalFailure extends RuntimeException {
  private static final long serialVersionUID = 1L;

  public FakeEvalFailure(String message) {
    super(message);
  }
}
