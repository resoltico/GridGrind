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
  /** Returns the effective default print-layout payload for one unconfigured sheet. */
  public static PrintLayoutInput defaults() {
    return new PrintLayoutInput(
        defaultPrintArea(),
        defaultOrientation(),
        defaultScaling(),
        defaultRepeatingRows(),
        defaultRepeatingColumns(),
        defaultHeader(),
        defaultFooter(),
        PrintSetupInput.defaults());
  }

  /** Creates a print-layout payload with the documented default advanced setup block. */
  public static PrintLayoutInput withDefaultSetup(
      PrintAreaInput printArea,
      ExcelPrintOrientation orientation,
      PrintScalingInput scaling,
      PrintTitleRowsInput repeatingRows,
      PrintTitleColumnsInput repeatingColumns,
      HeaderFooterTextInput header,
      HeaderFooterTextInput footer) {
    return new PrintLayoutInput(
        printArea,
        orientation,
        scaling,
        repeatingRows,
        repeatingColumns,
        header,
        footer,
        PrintSetupInput.defaults());
  }

  public PrintLayoutInput {
    Objects.requireNonNull(printArea, "printArea must not be null");
    Objects.requireNonNull(orientation, "orientation must not be null");
    Objects.requireNonNull(scaling, "scaling must not be null");
    Objects.requireNonNull(repeatingRows, "repeatingRows must not be null");
    Objects.requireNonNull(repeatingColumns, "repeatingColumns must not be null");
    Objects.requireNonNull(header, "header must not be null");
    Objects.requireNonNull(footer, "footer must not be null");
    Objects.requireNonNull(setup, "setup must not be null");
  }

  private static PrintAreaInput defaultPrintArea() {
    return new PrintAreaInput.None();
  }

  private static ExcelPrintOrientation defaultOrientation() {
    return ExcelPrintOrientation.PORTRAIT;
  }

  private static PrintScalingInput defaultScaling() {
    return new PrintScalingInput.Automatic();
  }

  private static PrintTitleRowsInput defaultRepeatingRows() {
    return new PrintTitleRowsInput.None();
  }

  private static PrintTitleColumnsInput defaultRepeatingColumns() {
    return new PrintTitleColumnsInput.None();
  }

  private static HeaderFooterTextInput defaultHeader() {
    return HeaderFooterTextInput.blank();
  }

  private static HeaderFooterTextInput defaultFooter() {
    return HeaderFooterTextInput.blank();
  }
}
