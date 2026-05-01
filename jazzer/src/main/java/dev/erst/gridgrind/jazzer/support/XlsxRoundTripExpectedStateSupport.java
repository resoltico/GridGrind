package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelChartSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelDataValidationSnapshot;
import dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPivotTableSnapshot;
import dev.erst.gridgrind.excel.ExcelRichTextSnapshot;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.util.CellReference;

/** Captures the pre-save workbook state that must survive `.xlsx` round-trips. */
final class XlsxRoundTripExpectedStateSupport {
  private XlsxRoundTripExpectedStateSupport() {}

  static ExpectedWorkbookState expectedWorkbookState(
      ExcelWorkbook workbook, List<WorkbookCommand> commands) throws IOException {
    ExpectedWorkbookFootprint footprint =
        XlsxRoundTripExpectedFootprintSupport.expectedWorkbookFootprint(commands);
    Map<String, List<ExcelCellSnapshot>> candidateSnapshots =
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

  static dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary expectedWorkbookSummary(
      ExcelWorkbook workbook) {
    return XlsxRoundTripExpectedWorkbookSurfaceSupport.expectedWorkbookSummary(workbook);
  }

  record ExpectedWorkbookState(
      Map<String, Map<CellCoordinate, ExcelCellStyleSnapshot>> expectedStyles,
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadata,
      Map<String, Map<CellCoordinate, ExcelRichTextSnapshot>> expectedRichText,
      Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges,
      dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary expectedWorkbookSummary,
      Map<String, dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary> expectedSheetSummaries,
      Map<String, ExpectedSheetLayoutState> expectedSheetLayouts,
      Map<String, List<ExcelDataValidationSnapshot>> expectedDataValidations,
      Map<String, List<ExcelConditionalFormattingBlockSnapshot>> expectedConditionalFormatting,
      Map<String, List<ExcelAutofilterSnapshot>> expectedAutofilters,
      Map<String, List<ExcelDrawingObjectSnapshot>> expectedDrawingObjects,
      Map<String, List<ExcelChartSnapshot>> expectedCharts,
      List<ExcelPivotTableSnapshot> expectedPivots,
      List<ExcelTableSnapshot> expectedTables) {}

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

  record ExpectedCellMetadata(ExcelHyperlink hyperlink, ExcelComment comment) {
    static ExpectedCellMetadata from(ExcelCellMetadataSnapshot metadata) {
      return new ExpectedCellMetadata(
          metadata.hyperlink().orElse(null),
          metadata.comment().map(derivedComment -> derivedComment.toPlainComment()).orElse(null));
    }
  }

  record ExpectedSheetLayoutState(
      ExcelSheetPane pane,
      Integer zoomPercent,
      Map<Integer, ExpectedRowLayoutState> rows,
      Map<Integer, ExpectedColumnLayoutState> columns) {}

  record ExpectedRowLayoutState(Boolean hidden, Integer outlineLevel, Boolean collapsed) {}

  record ExpectedColumnLayoutState(Boolean hidden, Integer outlineLevel, Boolean collapsed) {}

  record NamedRangeKey(String name, ExcelNamedRangeScope scope) {
    String displayName() {
      return switch (scope) {
        case ExcelNamedRangeScope.WorkbookScope _ -> "WORKBOOK:" + name;
        case ExcelNamedRangeScope.SheetScope sheetScope ->
            "SHEET:" + sheetScope.sheetName() + ":" + name;
      };
    }
  }

  record ExpectedNamedRange(
      ExcelNamedRangeScope scope, String refersToFormula, ExcelNamedRangeTarget target) {}
}
