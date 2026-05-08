package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Workbook-core chart selection used by chart introspection reads. */
public sealed interface ExcelChartSelection
    permits ExcelChartSelection.AllOnSheet, ExcelChartSelection.ByName {

  /** Owning sheet that scopes the selection. */
  String sheetName();

  /** Selects every chart on one sheet. */
  record AllOnSheet(String sheetName) implements ExcelChartSelection {
    public AllOnSheet {
      sheetName = requireNonBlank(sheetName, "sheetName");
    }
  }

  /** Selects one chart by its sheet-local name. */
  record ByName(String sheetName, String chartName) implements ExcelChartSelection {
    public ByName {
      sheetName = requireNonBlank(sheetName, "sheetName");
      chartName = requireNonBlank(chartName, "chartName");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
