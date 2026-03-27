package dev.erst.gridgrind.protocol;

import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;

/** Protocol-facing explicit cell or rectangular range target for named-range authoring. */
public record NamedRangeTarget(String sheetName, String range) {
  public NamedRangeTarget {
    ExcelNamedRangeTarget normalizedTarget = new ExcelNamedRangeTarget(sheetName, range);
    sheetName = normalizedTarget.sheetName();
    range = normalizedTarget.range();
  }

  /** Converts this transport target into the workbook-core representation. */
  public ExcelNamedRangeTarget toExcelNamedRangeTarget() {
    return new ExcelNamedRangeTarget(sheetName, range);
  }
}
