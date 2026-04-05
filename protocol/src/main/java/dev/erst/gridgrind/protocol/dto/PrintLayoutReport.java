package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import java.util.Objects;

/** Supported print-layout facts for one sheet. */
public record PrintLayoutReport(
    String sheetName,
    PrintAreaReport printArea,
    ExcelPrintOrientation orientation,
    PrintScalingReport scaling,
    PrintTitleRowsReport repeatingRows,
    PrintTitleColumnsReport repeatingColumns,
    HeaderFooterTextReport header,
    HeaderFooterTextReport footer) {
  public PrintLayoutReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(printArea, "printArea must not be null");
    Objects.requireNonNull(orientation, "orientation must not be null");
    Objects.requireNonNull(scaling, "scaling must not be null");
    Objects.requireNonNull(repeatingRows, "repeatingRows must not be null");
    Objects.requireNonNull(repeatingColumns, "repeatingColumns must not be null");
    Objects.requireNonNull(header, "header must not be null");
    Objects.requireNonNull(footer, "footer must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
  }
}
