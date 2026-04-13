package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.util.CellRangeAddress;

/** One ignored-error block anchored to one normalized A1-style sheet range. */
public record ExcelIgnoredError(String range, List<ExcelIgnoredErrorType> errorTypes) {
  public ExcelIgnoredError {
    range = normalizeRange(range);
    errorTypes = copyTypes(errorTypes);
    if (errorTypes.isEmpty()) {
      throw new IllegalArgumentException("errorTypes must not be empty");
    }
  }

  private static String normalizeRange(String range) {
    Objects.requireNonNull(range, "range must not be null");
    ExcelRange parsed = ExcelRange.parse(range);
    return new CellRangeAddress(
            parsed.firstRow(), parsed.lastRow(), parsed.firstColumn(), parsed.lastColumn())
        .formatAsString();
  }

  private static List<ExcelIgnoredErrorType> copyTypes(List<ExcelIgnoredErrorType> errorTypes) {
    List<ExcelIgnoredErrorType> copy =
        List.copyOf(Objects.requireNonNull(errorTypes, "errorTypes must not be null"));
    if (copy.size() != new LinkedHashSet<>(copy).size()) {
      throw new IllegalArgumentException("errorTypes must not contain duplicates");
    }
    return copy;
  }
}
