package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import java.util.Objects;

/** One authoritative supported print-layout payload in protocol form. */
public record PrintLayoutInput(
    PrintAreaInput printArea,
    ExcelPrintOrientation orientation,
    PrintScalingInput scaling,
    PrintTitleRowsInput repeatingRows,
    PrintTitleColumnsInput repeatingColumns,
    HeaderFooterTextInput header,
    HeaderFooterTextInput footer,
    PrintSetupInput setup) {
  /** Creates a print-layout payload while defaulting the advanced setup block. */
  public PrintLayoutInput(
      PrintAreaInput printArea,
      ExcelPrintOrientation orientation,
      PrintScalingInput scaling,
      PrintTitleRowsInput repeatingRows,
      PrintTitleColumnsInput repeatingColumns,
      HeaderFooterTextInput header,
      HeaderFooterTextInput footer) {
    this(printArea, orientation, scaling, repeatingRows, repeatingColumns, header, footer, null);
  }

  public PrintLayoutInput {
    printArea = printArea == null ? new PrintAreaInput.None() : printArea;
    orientation = orientation == null ? ExcelPrintOrientation.PORTRAIT : orientation;
    scaling = scaling == null ? new PrintScalingInput.Automatic() : scaling;
    repeatingRows = repeatingRows == null ? new PrintTitleRowsInput.None() : repeatingRows;
    repeatingColumns =
        repeatingColumns == null ? new PrintTitleColumnsInput.None() : repeatingColumns;
    header = header == null ? HeaderFooterTextInput.blank() : header;
    footer = footer == null ? HeaderFooterTextInput.blank() : footer;
    setup = setup == null ? PrintSetupInput.defaults() : setup;
    Objects.requireNonNull(printArea, "printArea must not be null");
    Objects.requireNonNull(orientation, "orientation must not be null");
    Objects.requireNonNull(scaling, "scaling must not be null");
    Objects.requireNonNull(repeatingRows, "repeatingRows must not be null");
    Objects.requireNonNull(repeatingColumns, "repeatingColumns must not be null");
    Objects.requireNonNull(header, "header must not be null");
    Objects.requireNonNull(footer, "footer must not be null");
    Objects.requireNonNull(setup, "setup must not be null");
  }
}
