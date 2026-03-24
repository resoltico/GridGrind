package org.apache.poi.ss.formula.eval;

/**
 * Test double for POI NotImplementedFunctionException to exercise unsupported formula detection.
 */
public final class FakeNotImplementedFunctionException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  public FakeNotImplementedFunctionException(String message) {
    super(message);
  }
}
