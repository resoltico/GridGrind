package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;
import java.util.Optional;

/** Known native-content coordinates, without inventing absent parser or package facts. */
public record WorkbookContentLocation(
    Optional<SheetLabel> sheetName,
    Optional<String> partName,
    Optional<String> fieldName,
    Optional<Integer> line,
    Optional<Integer> column) {
  public WorkbookContentLocation {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    partName = nonblank(partName, "partName");
    fieldName = nonblank(fieldName, "fieldName");
    line = positive(line, "line");
    column = positive(column, "column");
    if (sheetName.isEmpty()
        && partName.isEmpty()
        && fieldName.isEmpty()
        && line.isEmpty()
        && column.isEmpty()) {
      throw new IllegalArgumentException("Native content location must contain a known coordinate");
    }
  }

  private static Optional<String> nonblank(Optional<String> value, String field) {
    Objects.requireNonNull(value, field + " must not be null");
    value.ifPresent(text -> TextValidation.requireNonBlank(text, field));
    return value;
  }

  private static Optional<Integer> positive(Optional<Integer> value, String field) {
    Objects.requireNonNull(value, field + " must not be null");
    value.ifPresent(number -> {
      if (number < 1) throw new IllegalArgumentException(field + " must be positive");
    });
    return value;
  }
}
