package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import java.util.Objects;
import org.apache.poi.ss.usermodel.IgnoredErrorType;

/** Maps ignored-error enums between GridGrind and Apache POI. */
final class ExcelIgnoredErrorPoiBridge {
  private ExcelIgnoredErrorPoiBridge() {}

  static IgnoredErrorType toPoi(ExcelIgnoredErrorType ignoredErrorType) {
    Objects.requireNonNull(ignoredErrorType, "ignoredErrorType must not be null");
    return IgnoredErrorType.valueOf(ignoredErrorType.name());
  }

  static ExcelIgnoredErrorType fromPoi(IgnoredErrorType ignoredErrorType) {
    Objects.requireNonNull(ignoredErrorType, "ignoredErrorType must not be null");
    return ExcelIgnoredErrorType.valueOf(ignoredErrorType.name());
  }
}
