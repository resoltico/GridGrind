package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.PaneReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.apache.poi.ss.util.CellReference;
import org.jspecify.annotations.Nullable;

/** Captures the pre-save workbook state that must survive `.xlsx` round-trips. */
final class XlsxRoundTripExpectedStateSupport {
  private XlsxRoundTripExpectedStateSupport() {}

  static ExpectedWorkbookState expectedWorkbookState(
      ExcelWorkbook workbook, List<ResolvedStep> commands) throws IOException {
    ExpectedWorkbookFootprint footprint =
        XlsxRoundTripExpectedFootprintSupport.expectedWorkbookFootprint(commands);
    Map<String, List<CellReport>> candidateSnapshots =
        XlsxRoundTripExpectedFootprintSupport.expectedCellSnapshots(workbook, footprint);
    return new ExpectedWorkbookState(
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedStyles(
            candidateSnapshots, XlsxRoundTripVerifier.defaultStyleSnapshot()),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedMetadata(candidateSnapshots),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedRichText(candidateSnapshots),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedNamedRanges(workbook),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedWorkbookSummary(workbook),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedSheetSummaries(workbook),
        XlsxRoundTripExpectedLayoutSupport.expectedSheetLayouts(workbook),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedDataValidations(workbook),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedConditionalFormatting(workbook),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedAutofilters(workbook),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedDrawingObjects(workbook),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedCharts(workbook),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedPivots(workbook),
        XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedTables(workbook));
  }

  static dev.erst.gridgrind.core.result.WorkbookSummary expectedWorkbookSummary(
      ExcelWorkbook workbook) {
    return XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedWorkbookSummary(workbook);
  }

  record ExpectedWorkbookState(
      Map<String, Map<CellCoordinate, CellStyleReport>> expectedStyles,
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadata,
      Map<String, Map<CellCoordinate, List<RichTextRunReport>>> expectedRichText,
      Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges,
      dev.erst.gridgrind.core.result.WorkbookSummary expectedWorkbookSummary,
      Map<String, dev.erst.gridgrind.core.result.SheetSummaryReport> expectedSheetSummaries,
      Map<String, ExpectedSheetLayoutState> expectedSheetLayouts,
      Map<String, List<DataValidationEntryReport>> expectedDataValidations,
      Map<String, List<ConditionalFormattingEntryReport>> expectedConditionalFormatting,
      Map<String, List<AutofilterEntryReport>> expectedAutofilters,
      Map<String, List<DrawingObjectReport>> expectedDrawingObjects,
      Map<String, List<ChartReport>> expectedCharts,
      List<PivotTableReport> expectedPivots,
      List<TableEntryReport> expectedTables) {}

  record ExpectedWorkbookFootprint(Map<String, List<CellCoordinate>> candidateCoordinatesBySheet) {}

  record CellCoordinate(int rowIndex, int columnIndex) {
    static CellCoordinate fromAddress(String address) {
      CellReference cellReference = new CellReference(address);
      return new CellCoordinate(cellReference.getRow(), cellReference.getCol());
    }

    String a1Address() {
      return new CellReference(rowIndex, columnIndex).formatAsString();
    }
  }

  record ExpectedCellMetadata(
      Optional<HyperlinkTarget> hyperlink, Optional<ResolvedComment> comment) {
    static ExpectedCellMetadata from(CellReport metadata) {
      return new ExpectedCellMetadata(
          metadata.hyperlink(),
          metadata
              .comment()
              .map(
                  comment ->
                      new ResolvedComment(comment.text(), comment.author(), comment.visible())));
    }
  }

  record ExpectedSheetLayoutState(
      PaneReport pane,
      Integer zoomPercent,
      Map<Integer, ExpectedRowLayoutState> rows,
      Map<Integer, ExpectedColumnLayoutState> columns) {}

  record ExpectedRowLayoutState(Boolean hidden, Integer outlineLevel, Boolean collapsed) {}

  record ExpectedColumnLayoutState(Boolean hidden, Integer outlineLevel, Boolean collapsed) {}

  record NamedRangeKey(String name, NamedRangeScopeReport scope) {
    String displayName() {
      return switch (scope) {
        case NamedRangeScopeReport.Workbook _ -> "WORKBOOK:" + name;
        case NamedRangeScopeReport.Sheet sheetScope ->
            "SHEET:" + sheetScope.sheetName() + ":" + name;
      };
    }
  }

  record ExpectedNamedRange(
      NamedRangeScopeReport scope, String refersToFormula, @Nullable NamedRangeTarget target) {}
}
