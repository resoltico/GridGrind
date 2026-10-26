package dev.erst.gridgrind.core.model.print;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import java.util.Objects;

/** One authoritative supported print-layout payload in protocol form. */
@CatalogType(
    id = "PrintLayoutSpec",
    summary =
        "Authoritative supported print-layout payload for one SET_PRINT_LAYOUT request. Supply explicit printArea, orientation, scaling, repeatingRows, repeatingColumns, header, footer, and setup fields.",
    displayOrder = 63)
public record PrintLayoutSpec(
    PrintAreaSpec printArea,
    ExcelPrintOrientation orientation,
    PrintScalingSpec scaling,
    PrintTitleRowsSpec repeatingRows,
    PrintTitleColumnsSpec repeatingColumns,
    HeaderFooterTextSpec header,
    HeaderFooterTextSpec footer,
    PrintSetupSpec setup) {
  /** Returns the effective default print-layout payload for one unconfigured sheet. */
  public static PrintLayoutSpec defaults() {
    return new PrintLayoutSpec(
        defaultPrintArea(),
        defaultOrientation(),
        defaultScaling(),
        defaultRepeatingRows(),
        defaultRepeatingColumns(),
        defaultHeader(),
        defaultFooter(),
        PrintSetupSpec.defaults());
  }

  /** Creates a print-layout payload with the documented default advanced setup block. */
  public static PrintLayoutSpec withDefaultSetup(
      PrintAreaSpec printArea,
      ExcelPrintOrientation orientation,
      PrintScalingSpec scaling,
      PrintTitleRowsSpec repeatingRows,
      PrintTitleColumnsSpec repeatingColumns,
      HeaderFooterTextSpec header,
      HeaderFooterTextSpec footer) {
    return new PrintLayoutSpec(
        printArea,
        orientation,
        scaling,
        repeatingRows,
        repeatingColumns,
        header,
        footer,
        PrintSetupSpec.defaults());
  }

  public PrintLayoutSpec {
    Objects.requireNonNull(printArea, "printArea must not be null");
    Objects.requireNonNull(orientation, "orientation must not be null");
    Objects.requireNonNull(scaling, "scaling must not be null");
    Objects.requireNonNull(repeatingRows, "repeatingRows must not be null");
    Objects.requireNonNull(repeatingColumns, "repeatingColumns must not be null");
    Objects.requireNonNull(header, "header must not be null");
    Objects.requireNonNull(footer, "footer must not be null");
    Objects.requireNonNull(setup, "setup must not be null");
  }

  private static PrintAreaSpec defaultPrintArea() {
    return new PrintAreaSpec.None();
  }

  private static ExcelPrintOrientation defaultOrientation() {
    return ExcelPrintOrientation.PORTRAIT;
  }

  private static PrintScalingSpec defaultScaling() {
    return new PrintScalingSpec.Automatic();
  }

  private static PrintTitleRowsSpec defaultRepeatingRows() {
    return new PrintTitleRowsSpec.None();
  }

  private static PrintTitleColumnsSpec defaultRepeatingColumns() {
    return new PrintTitleColumnsSpec.None();
  }

  private static HeaderFooterTextSpec defaultHeader() {
    return HeaderFooterTextSpec.blank();
  }

  private static HeaderFooterTextSpec defaultFooter() {
    return HeaderFooterTextSpec.blank();
  }
}
