package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import java.util.Objects;
import org.apache.poi.ss.usermodel.PaneType;

/** Maps pane-region enums between GridGrind and Apache POI. */
final class ExcelPanePoiBridge {
  private ExcelPanePoiBridge() {}

  static ExcelPaneRegion fromPoi(PaneType paneType) {
    Objects.requireNonNull(paneType, "paneType must not be null");
    return ExcelPaneRegion.valueOf(paneType.name());
  }

  static PaneType toPoi(ExcelPaneRegion paneRegion) {
    Objects.requireNonNull(paneRegion, "paneRegion must not be null");
    return PaneType.valueOf(paneRegion.name());
  }
}
