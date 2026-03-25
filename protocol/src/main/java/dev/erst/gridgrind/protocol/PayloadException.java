package dev.erst.gridgrind.protocol;

/**
 * Sealed marker interface for JSON payload failures with structured path and location metadata.
 * Each permitted subtype is a concrete exception class carrying the JSON path and line/column
 * coordinates directly.
 */
public sealed interface PayloadException permits InvalidJsonException, InvalidRequestException {
  /** JSON pointer path to the element that triggered the failure. */
  String jsonPath();

  /** Line number within the JSON payload where the error was detected. */
  Integer jsonLine();

  /** Column number within the JSON payload where the error was detected. */
  Integer jsonColumn();
}
