package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing alignment patch used by {@link CellStyleInput}. */
public record CellAlignmentInput(
    Optional<Boolean> wrapText,
    Optional<ExcelHorizontalAlignment> horizontalAlignment,
    Optional<ExcelVerticalAlignment> verticalAlignment,
    Optional<Integer> textRotation,
    Optional<Integer> indentation) {
  private static final int MAX_TEXT_ROTATION = 180;
  private static final int MAX_INDENTATION = 250;

  public CellAlignmentInput {
    Objects.requireNonNull(wrapText, "wrapText must not be null");
    Objects.requireNonNull(horizontalAlignment, "horizontalAlignment must not be null");
    Objects.requireNonNull(verticalAlignment, "verticalAlignment must not be null");
    Objects.requireNonNull(textRotation, "textRotation must not be null");
    Objects.requireNonNull(indentation, "indentation must not be null");
    if (wrapText.isEmpty()
        && horizontalAlignment.isEmpty()
        && verticalAlignment.isEmpty()
        && textRotation.isEmpty()
        && indentation.isEmpty()) {
      throw new IllegalArgumentException("alignment must set at least one attribute");
    }
    if (textRotation.isPresent()
        && (textRotation.orElseThrow() < 0 || textRotation.orElseThrow() > MAX_TEXT_ROTATION)) {
      throw new IllegalArgumentException(
          "textRotation must be between 0 and " + MAX_TEXT_ROTATION + " (inclusive)");
    }
    if (indentation.isPresent()
        && (indentation.orElseThrow() < 0 || indentation.orElseThrow() > MAX_INDENTATION)) {
      throw new IllegalArgumentException(
          "indentation must be between 0 and " + MAX_INDENTATION + " (inclusive)");
    }
  }
}
