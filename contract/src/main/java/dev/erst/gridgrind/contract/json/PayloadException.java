package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/**
 * Sealed marker interface for JSON payload failures with structured path and location metadata.
 * Each permitted subtype is a concrete exception class carrying the JSON path and line/column
 * coordinates directly.
 */
public sealed interface PayloadException
    permits InvalidJsonException, InvalidRequestException, InvalidRequestShapeException {
  /** JSON pointer path to the element that triggered the failure. */
  Optional<String> jsonPath();

  /** Line number within the JSON payload where the error was detected. */
  Optional<Integer> jsonLine();

  /** Column number within the JSON payload where the error was detected. */
  Optional<Integer> jsonColumn();
}
