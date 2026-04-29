package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
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
        new PrintAreaInput.None(),
        ExcelPrintOrientation.PORTRAIT,
        new PrintScalingInput.Automatic(),
        new PrintTitleRowsInput.None(),
        new PrintTitleColumnsInput.None(),
        HeaderFooterTextInput.blank(),
        HeaderFooterTextInput.blank(),
        PrintSetupInput.defaults());
  }

  /** Creates a print-layout payload while defaulting the advanced setup block. */
  public PrintLayoutInput(
      PrintAreaInput printArea,
      ExcelPrintOrientation orientation,
      PrintScalingInput scaling,
      PrintTitleRowsInput repeatingRows,
      PrintTitleColumnsInput repeatingColumns,
      HeaderFooterTextInput header,
      HeaderFooterTextInput footer) {
    this(
        printArea,
        orientation,
        scaling,
        repeatingRows,
        repeatingColumns,
        header,
        footer,
        PrintSetupInput.defaults());
  }

  /** Creates one layout while defaulting the advanced setup block. */
  public static PrintLayoutInput create(
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

  @JsonCreator(mode = JsonCreator.Mode.DELEGATING)
  static PrintLayoutInput create(PrintLayoutJson json) {
    PrintLayoutInput defaults = defaults();
    return new PrintLayoutInput(
        defaultValue(json.printArea(), defaults.printArea()),
        defaultValue(json.orientation(), defaults.orientation()),
        defaultValue(json.scaling(), defaults.scaling()),
        defaultValue(json.repeatingRows(), defaults.repeatingRows()),
        defaultValue(json.repeatingColumns(), defaults.repeatingColumns()),
        defaultValue(json.header(), defaults.header()),
        defaultValue(json.footer(), defaults.footer()),
        defaultValue(json.setup(), defaults.setup()));
  }

  private static <T> T defaultValue(T value, T defaultValue) {
    return value == null ? defaultValue : value;
  }

  private record PrintLayoutJson(
      PrintAreaInput printArea,
      ExcelPrintOrientation orientation,
      PrintScalingInput scaling,
      PrintTitleRowsInput repeatingRows,
      PrintTitleColumnsInput repeatingColumns,
      HeaderFooterTextInput header,
      HeaderFooterTextInput footer,
      PrintSetupInput setup) {}
}
