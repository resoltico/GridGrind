package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** One authoritative supported print-layout payload in protocol form. */
public record PrintLayoutInput(
    PrintAreaInput printArea,
    PrintOrientation orientation,
    PrintScalingInput scaling,
    PrintTitleRowsInput repeatingRows,
    PrintTitleColumnsInput repeatingColumns,
    HeaderFooterTextInput header,
    HeaderFooterTextInput footer) {
  public PrintLayoutInput {
    printArea = printArea == null ? new PrintAreaInput.None() : printArea;
    orientation = orientation == null ? PrintOrientation.PORTRAIT : orientation;
    scaling = scaling == null ? new PrintScalingInput.Automatic() : scaling;
    repeatingRows = repeatingRows == null ? new PrintTitleRowsInput.None() : repeatingRows;
    repeatingColumns =
        repeatingColumns == null ? new PrintTitleColumnsInput.None() : repeatingColumns;
    header = header == null ? HeaderFooterTextInput.blank() : header;
    footer = footer == null ? HeaderFooterTextInput.blank() : footer;
    Objects.requireNonNull(printArea, "printArea must not be null");
    Objects.requireNonNull(orientation, "orientation must not be null");
    Objects.requireNonNull(scaling, "scaling must not be null");
    Objects.requireNonNull(repeatingRows, "repeatingRows must not be null");
    Objects.requireNonNull(repeatingColumns, "repeatingColumns must not be null");
    Objects.requireNonNull(header, "header must not be null");
    Objects.requireNonNull(footer, "footer must not be null");
  }
}
