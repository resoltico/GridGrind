package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CellInput;
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
      ExcelCellSnapshot snapshot, CellInput expectedValue, boolean date1904) {
    dev.erst.gridgrind.contract.dto.CellReport report =
        InspectionResultCellReportSupport.toCellReport(snapshot, KEY_MATCH_PROJECTION, date1904);
    return switch (expectedValue) {
      case CellInput.Blank _ ->
          report instanceof dev.erst.gridgrind.contract.dto.CellReport.BlankReport;
      case CellInput.Text text ->
          report instanceof dev.erst.gridgrind.contract.dto.CellReport.TextReport textReport
              && textReport
                  .textValue()
                  .orElseThrow()
                  .equals(inlineText(text.source(), "table row key TEXT"));
      case CellInput.NumberValue numberValue ->
          report instanceof dev.erst.gridgrind.contract.dto.CellReport.NumberReport numberReport
              && Double.compare(numberReport.numberValue().orElseThrow(), numberValue.number())
                  == 0;
      case CellInput.BooleanValue booleanValue ->
          report instanceof dev.erst.gridgrind.contract.dto.CellReport.BooleanReport booleanReport
              && booleanReport.booleanValue().orElseThrow().equals(booleanValue.bool());
      case CellInput.Formula formula ->
          report instanceof dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formulaReport
              && formulaReport
                  .formula()
                  .orElseThrow()
                  .equals(inlineText(formula.source(), "table row key FORMULA"));
      default -> false;
    };
  }

  private static String inlineText(
      dev.erst.gridgrind.contract.source.TextSourceInput source, String context) {
    if (source instanceof dev.erst.gridgrind.contract.source.TextSourceInput.Inline inline) {
      return inline.text();
    }
    throw new IllegalStateException(context + " must be resolved to INLINE text before execution");
  }
}
