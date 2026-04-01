package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Read-time view of one data-validation structure loaded from a sheet. */
public sealed interface ExcelDataValidationSnapshot
    permits ExcelDataValidationSnapshot.Supported, ExcelDataValidationSnapshot.Unsupported {

  /** Covered A1-style ranges stored on the sheet for this validation structure. */
  List<String> ranges();

  /** Fully supported data-validation structure. */
  record Supported(List<String> ranges, ExcelDataValidationDefinition validation)
      implements ExcelDataValidationSnapshot {
    public Supported {
      ranges = copyRanges(ranges);
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Present workbook structure that GridGrind can detect but not yet model fully. */
  record Unsupported(List<String> ranges, String kind, String detail)
      implements ExcelDataValidationSnapshot {
    public Unsupported {
      ranges = copyRanges(ranges);
      kind = requireNonBlank(kind, "kind");
      detail = requireNonBlank(detail, "detail");
    }
  }

  private static List<String> copyRanges(List<String> ranges) {
    Objects.requireNonNull(ranges, "ranges must not be null");
    List<String> copy = List.copyOf(ranges);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("ranges must not be empty");
    }
    for (String range : copy) {
      requireNonBlank(range, "ranges");
    }
    return copy;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
