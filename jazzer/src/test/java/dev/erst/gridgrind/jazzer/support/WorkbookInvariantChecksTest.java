package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelHyperlinkType;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.protocol.CellSelection;
import dev.erst.gridgrind.protocol.FontHeightReport;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindProtocolVersion;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.NamedRangeScope;
import dev.erst.gridgrind.protocol.NamedRangeSelection;
import dev.erst.gridgrind.protocol.NamedRangeTarget;
import dev.erst.gridgrind.protocol.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.WorkbookReadResult;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests for WorkbookInvariantChecks success-shape validation. */
class WorkbookInvariantChecksTest {
  @Test
  void acceptsSuccessResponsesWithOrderedReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");
    GridGrindResponse.CellStyleReport style = defaultStyle();

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.Saved(workbookPath.toString()),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 1, false)),
                new WorkbookReadResult.NamedRangesResult(
                    "ranges",
                    List.of(
                        new GridGrindResponse.NamedRangeReport.RangeReport(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            "Budget!$B$4",
                            new NamedRangeTarget("Budget", "B4")))),
                new WorkbookReadResult.CellsResult(
                    "cells",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellReport.TextReport(
                            "A1",
                            "STRING",
                            "Report",
                            style,
                            new GridGrindResponse.HyperlinkReport(
                                ExcelHyperlinkType.URL, "https://example.com/report"),
                            new GridGrindResponse.CommentReport("Review", "GridGrind", true),
                            "Report"))),
                new WorkbookReadResult.WindowResult(
                    "window",
                    new GridGrindResponse.WindowReport(
                        "Budget",
                        "A1",
                        1,
                        1,
                        List.of(
                            new GridGrindResponse.WindowRowReport(
                                0,
                                List.of(
                                    new GridGrindResponse.CellReport.TextReport(
                                        "A1",
                                        "STRING",
                                        "Report",
                                        style,
                                        null,
                                        null,
                                        "Report")))))),
                new WorkbookReadResult.MergedRegionsResult(
                    "merged",
                    "Budget",
                    List.of(new GridGrindResponse.MergedRegionReport("A1:B1"))),
                new WorkbookReadResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1",
                            new GridGrindResponse.HyperlinkReport(
                                ExcelHyperlinkType.URL, "https://example.com/report")))),
                new WorkbookReadResult.CommentsResult(
                    "comments",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellCommentReport(
                            "A1",
                            new GridGrindResponse.CommentReport("Review", "GridGrind", true)))),
                new WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new GridGrindResponse.SheetLayoutReport(
                        "Budget",
                        new GridGrindResponse.FreezePaneReport.Frozen(1, 1, 1, 1),
                        List.of(new GridGrindResponse.ColumnLayoutReport(0, 12.5)),
                        List.of(new GridGrindResponse.RowLayoutReport(0, 18.0)))),
                new WorkbookReadResult.FormulaSurfaceResult(
                    "formula-surface",
                    new GridGrindResponse.FormulaSurfaceReport(
                        1,
                        List.of(
                            new GridGrindResponse.SheetFormulaSurfaceReport(
                                "Budget",
                                1,
                                1,
                                List.of(
                                    new GridGrindResponse.FormulaPatternReport(
                                        "SUM(B2:B3)", 1, List.of("B4"))))))),
                new WorkbookReadResult.SheetSchemaResult(
                    "schema",
                    new GridGrindResponse.SheetSchemaReport(
                        "Budget",
                        "A1",
                        2,
                        1,
                        1,
                        List.of(
                            new GridGrindResponse.SchemaColumnReport(
                                0,
                                "A1",
                                "Item",
                                1,
                                0,
                                List.of(new GridGrindResponse.TypeCountReport("STRING", 1)),
                                "STRING")))),
                new WorkbookReadResult.NamedRangeSurfaceResult(
                    "named-range-surface",
                    new GridGrindResponse.NamedRangeSurfaceReport(
                        1,
                        0,
                        1,
                        0,
                        List.of(
                            new GridGrindResponse.NamedRangeSurfaceEntryReport(
                                "BudgetTotal",
                                new NamedRangeScope.Workbook(),
                                "Budget!$B$4",
                                GridGrindResponse.NamedRangeBackingKind.RANGE))))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForOrderedReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadOperation.GetWorkbookSummary("summary"),
                new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")),
                new WorkbookReadOperation.GetHyperlinks(
                    "hyperlinks", "Budget", new CellSelection.AllUsedCells()),
                new WorkbookReadOperation.AnalyzeNamedRangeSurface(
                    "surface", new NamedRangeSelection.All())));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.Saved(workbookPath.toString()),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 1, false)),
                new WorkbookReadResult.CellsResult(
                    "cells", "Budget", List.of(textCell("A1", "Report"))),
                new WorkbookReadResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1",
                            new GridGrindResponse.HyperlinkReport(
                                ExcelHyperlinkType.URL, "https://example.com/report")))),
                new WorkbookReadResult.NamedRangeSurfaceResult(
                    "surface",
                    new GridGrindResponse.NamedRangeSurfaceReport(
                        1,
                        0,
                        1,
                        0,
                        List.of(
                            new GridGrindResponse.NamedRangeSurfaceEntryReport(
                                "BudgetTotal",
                                new NamedRangeScope.Workbook(),
                                "Budget!$B$4",
                                GridGrindResponse.NamedRangeBackingKind.RANGE))))));

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

  private static GridGrindResponse.CellStyleReport defaultStyle() {
    return new GridGrindResponse.CellStyleReport(
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
  }

  private static GridGrindResponse.CellReport.TextReport textCell(String address, String value) {
    return new GridGrindResponse.CellReport.TextReport(
        address, "STRING", value, defaultStyle(), null, null, value);
  }
}
