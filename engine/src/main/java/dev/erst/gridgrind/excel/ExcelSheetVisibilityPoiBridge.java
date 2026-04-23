package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.util.Objects;
import org.apache.poi.ss.usermodel.SheetVisibility;

/** Maps sheet-visibility enums between GridGrind and Apache POI. */
final class ExcelSheetVisibilityPoiBridge {
  private ExcelSheetVisibilityPoiBridge() {}

  static ExcelSheetVisibility fromPoi(SheetVisibility visibility) {
    Objects.requireNonNull(visibility, "visibility must not be null");
    return ExcelSheetVisibility.valueOf(visibility.name());
  }

  static SheetVisibility toPoi(ExcelSheetVisibility visibility) {
    Objects.requireNonNull(visibility, "visibility must not be null");
    return SheetVisibility.valueOf(visibility.name());
  }
}
