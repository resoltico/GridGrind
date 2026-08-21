package dev.erst.gridgrind.contract.json;

import java.util.Objects;

/** Coordinates tolerant parsing with request-shape analysis. */
final class RequestStructuralAnalyzer {
  private RequestStructuralAnalyzer() {}

  static RequestAnalysis analyze(byte[] bytes) {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return RequestRootStructuralAnalyzer.analyze(TolerantRequestJsonParser.parse(bytes));
  }
}
