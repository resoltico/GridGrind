package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelHyperlinkType;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.protocol.FontHeightReport;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindProtocolVersion;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.NamedRangeScope;
import dev.erst.gridgrind.protocol.NamedRangeTarget;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookInvariantChecks success-shape validation. */
class WorkbookInvariantChecksTest {
  @Test
  void acceptsSuccessResponsesWithNamedRangesAndCellMetadata() {
    GridGrindResponse.CellStyleReport style =
        new GridGrindResponse.CellStyleReport(
            "General",
            false,
            false,
            false,
            ExcelHorizontalAlignment.GENERAL,
            ExcelVerticalAlignment.BOTTOM,
            "Calibri",
            new FontHeightReport(220, new BigDecimal("11")),
            null,
            false,
            false,
            null,
            ExcelBorderStyle.NONE,
            ExcelBorderStyle.NONE,
            ExcelBorderStyle.NONE,
            ExcelBorderStyle.NONE);
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            null,
            new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 2, false),
            List.of(
                new GridGrindResponse.NamedRangeReport.RangeReport(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    "Budget!$B$4",
                    new NamedRangeTarget("Budget", "B4"))),
            List.of(
                new GridGrindResponse.SheetReport(
                    "Budget",
                    1,
                    0,
                    0,
                    List.of(
                        new GridGrindResponse.CellReport.TextReport(
                            "A1",
                            "STRING",
                            "Report",
                            style,
                            new GridGrindResponse.HyperlinkReport(
                                ExcelHyperlinkType.URL, "https://example.com/report"),
                            new GridGrindResponse.CommentReport("Review", "GridGrind", true),
                            "Report")),
                    List.of())));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeWhenNamedRangeAnalysisIsNone() {
    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(),
            new GridGrindRequest.WorkbookAnalysisRequest(
                List.of(), new GridGrindRequest.WorkbookAnalysisRequest.NamedRangeInspection.None()));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            null,
            new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 1, false),
            List.of(),
            List.of());

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsWorkbookShapeWithNamedRanges() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("B4", ExcelCellValue.number(61.0));
      workbook.setNamedRange(
          new dev.erst.gridgrind.excel.ExcelNamedRangeDefinition(
              "BudgetTotal",
              new dev.erst.gridgrind.excel.ExcelNamedRangeScope.WorkbookScope(),
              new dev.erst.gridgrind.excel.ExcelNamedRangeTarget("Budget", "B4")));

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }
}
