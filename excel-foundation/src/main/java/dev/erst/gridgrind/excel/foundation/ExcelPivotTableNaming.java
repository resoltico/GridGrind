package dev.erst.gridgrind.excel.foundation;

import java.util.Objects;
import java.util.Optional;

/** Product-owned validation and normalization rules for persisted pivot-table names. */
public final class ExcelPivotTableNaming {
  private ExcelPivotTableNaming() {}

  /** Validates and returns one pivot-table name exactly as authored. */
  public static String validateName(String name) {
    Objects.requireNonNull(name, "name must not be null");
    violation(name)
        .ifPresent(
            violation -> {
              throw new IllegalArgumentException(renderViolation(violation));
            });
    return name;
  }

  /** Returns the first pivot-table naming violation, if any. */
  public static Optional<Violation> violation(String name) {
    Objects.requireNonNull(name, "name must not be null");
    return name.isBlank() ? Optional.of(Violation.BLANK) : Optional.empty();
  }

  private static String renderViolation(Violation violation) {
    return switch (violation) {
      case BLANK -> "name must not be blank";
    };
  }

  /** Structured pivot-table-name violation family. */
  public enum Violation {
    BLANK
  }
}
