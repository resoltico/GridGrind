package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/**
 * Sealed marker interface for JSON payload failures with structured path and location metadata.
 * Each permitted subtype is a concrete exception class carrying the JSON path and line/column
 * coordinates directly.
 */
public sealed interface PayloadException
    permits InvalidEncodingException,
        InvalidJsonException,
        InvalidRequestException,
        InvalidRequestShapeException,
        NumberNotRepresentableException {
  /** Normalized JSON location variant for the payload failure. */
  PayloadLocation jsonLocation();

  /** JSON pointer path to the element that triggered the failure. */
  default Optional<String> jsonPath() {
    return jsonLocation().jsonPath();
  }

  /** Line number within the JSON payload where the error was detected. */
  default Optional<Integer> jsonLine() {
    return jsonLocation().jsonLine();
  }

  /** Column number within the JSON payload where the error was detected. */
  default Optional<Integer> jsonColumn() {
    return jsonLocation().jsonColumn();
  }
}
