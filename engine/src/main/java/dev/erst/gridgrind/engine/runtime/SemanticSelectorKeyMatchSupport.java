package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.excel.ExcelCellReadFacet;
import dev.erst.gridgrind.excel.ExcelCellReadProjection;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import java.util.Set;

/** Shared cell-value matching for semantic table-row key resolution. */
final class SemanticSelectorKeyMatchSupport {
  private static final ExcelCellReadProjection KEY_MATCH_PROJECTION =
      new ExcelCellReadProjection(Set.of(ExcelCellReadFacet.VALUE, ExcelCellReadFacet.FORMULA));

  private SemanticSelectorKeyMatchSupport() {}

  static boolean matchesKeyCell(
      ExcelCellSnapshot snapshot, CellValueSpec expectedValue, boolean date1904) {
    dev.erst.gridgrind.core.model.value.CellReport report =
        InspectionResultCellReportSupport.toCellReport(snapshot, KEY_MATCH_PROJECTION, date1904);
    return switch (expectedValue) {
      case CellValueSpec.Blank _ ->
          report instanceof dev.erst.gridgrind.core.model.value.CellReport.BlankReport;
      case CellValueSpec.Text text ->
          report instanceof dev.erst.gridgrind.core.model.value.CellReport.TextReport textReport
              && textReport
                  .value()
                  .map(dev.erst.gridgrind.core.model.value.CellValue.Text::value)
                  .orElseThrow()
                  .equals(inlineText(text.value(), "table row key TEXT"));
      case CellValueSpec.NumberValue numberValue ->
          report instanceof dev.erst.gridgrind.core.model.value.CellReport.NumberReport numberReport
              && Double.compare(
                      numberReport
                          .value()
                          .map(dev.erst.gridgrind.core.model.value.CellValue.NumberValue::value)
                          .orElseThrow(),
                      numberValue.value())
                  == 0;
      case CellValueSpec.BooleanValue booleanValue ->
          report
                  instanceof
                  dev.erst.gridgrind.core.model.value.CellReport.BooleanReport booleanReport
              && booleanReport
                  .value()
                  .map(dev.erst.gridgrind.core.model.value.CellValue.BooleanValue::value)
                  .orElseThrow()
                  .equals(booleanValue.value());
      case CellValueSpec.Formula formula ->
          report
                  instanceof
                  dev.erst.gridgrind.core.model.value.CellReport.FormulaReport formulaReport
              && formulaReport
                  .formula()
                  .orElseThrow()
                  .equals(inlineText(formula.value(), "table row key FORMULA"));
      default -> false;
    };
  }

  private static String inlineText(
      dev.erst.gridgrind.core.source.TextSourceSpec source, String context) {
    if (source instanceof dev.erst.gridgrind.core.source.TextSourceSpec.Inline inline) {
      return inline.text();
    }
    throw new IllegalStateException(context + " must be resolved to INLINE text before execution");
  }
}
