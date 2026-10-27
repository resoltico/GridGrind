package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import java.util.Objects;

/** Print layout with materialized header/footer text and shared source-free settings. */
public record ResolvedPrintLayout(
    PrintAreaSpec printArea,
    ExcelPrintOrientation orientation,
    PrintScalingSpec scaling,
    PrintTitleRowsSpec repeatingRows,
    PrintTitleColumnsSpec repeatingColumns,
    HeaderFooterTextReport header,
    HeaderFooterTextReport footer,
    PrintSetupSpec setup) {
  /** Returns the GridGrind default print layout for a sheet with no explicit print settings. */
  public static ResolvedPrintLayout defaults() {
    return new ResolvedPrintLayout(
        new PrintAreaSpec.None(),
        ExcelPrintOrientation.PORTRAIT,
        new PrintScalingSpec.Automatic(),
        new PrintTitleRowsSpec.None(),
        new PrintTitleColumnsSpec.None(),
        HeaderFooterTextReport.blank(),
        HeaderFooterTextReport.blank(),
        PrintSetupSpec.defaults());
  }

  /** Creates a print layout while defaulting the advanced print-setup block. */
  public ResolvedPrintLayout(
      PrintAreaSpec printArea,
      ExcelPrintOrientation orientation,
      PrintScalingSpec scaling,
      PrintTitleRowsSpec repeatingRows,
      PrintTitleColumnsSpec repeatingColumns,
      HeaderFooterTextReport header,
      HeaderFooterTextReport footer) {
    this(
        printArea,
        orientation,
        scaling,
        repeatingRows,
        repeatingColumns,
        header,
        footer,
        PrintSetupSpec.defaults());
  }

  public ResolvedPrintLayout {
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
