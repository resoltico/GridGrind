package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** One structural defect found while analysing a request without stopping at the first defect. */
public sealed interface RequestStructuralProblem
    permits RequestInvalidEncoding,
        RequestInvalidJson,
        RequestDuplicateKey,
        RequestNumberNotRepresentable,
        RequestShapeStructuralProblem {

  /** Returns the JSON path when one unambiguously identifies the malformed value. */
  Optional<String> jsonPath();

  /** Returns the zero-based UTF-8 byte offset of the offending token when one exists. */
  Optional<Long> byteOffset();

  /** Returns the one-based JSON line when syntax analysis captured it. */
  default Optional<Integer> jsonLine() {
    return Optional.empty();
  }

  /** Returns the one-based JSON column when syntax analysis captured it. */
  default Optional<Integer> jsonColumn() {
    return Optional.empty();
  }

  /** Returns the stable, product-owned diagnostic message. */
  String message();
}
