package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelPaneRegion;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage tests for nested response-report records that are not exercised by end-to-end flows. */
class GridGrindResponseNestedCoverageTest {
  @Test
  void windowLayoutSchemaAndAnalysisReportsValidateAndCopyCollections() {
    GridGrindResponse.CellReport textCell =
        new GridGrindResponse.CellReport.TextReport(
            "A1",
            "STRING",
            "Owner",
            style(),
            null,
            null,
            "Owner",
            List.of(new RichTextRunReport("Owner", style().font())));
    GridGrindResponse.WindowRowReport row =
        new GridGrindResponse.WindowRowReport(0, List.of(textCell));
    GridGrindResponse.WindowReport window =
        new GridGrindResponse.WindowReport("Budget", "A1", 1, 1, List.of(row));
    assertEquals("Budget", window.sheetName());
    assertEquals(
        "rowCount must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.WindowReport("Budget", "A1", 0, 1, List.of(row)))
            .getMessage());

    GridGrindResponse.SheetLayoutReport layout =
        new GridGrindResponse.SheetLayoutReport(
            "Budget",
            new PaneReport.Split(1, 1, 0, 0, ExcelPaneRegion.LOWER_RIGHT),
            125,
            SheetPresentationReport.defaults(),
            List.of(new GridGrindResponse.ColumnLayoutReport(0, 8.43d, false, 0, false)),
            List.of(new GridGrindResponse.RowLayoutReport(0, 15.0d, false, 0, false)));
    assertEquals(125, layout.zoomPercent());
    assertEquals(
        "zoomPercent must be between 10 and 400 inclusive: 401",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.SheetLayoutReport(
                        "Budget",
                        new PaneReport.None(),
                        401,
                        SheetPresentationReport.defaults(),
                        List.of(),
                        List.of()))
            .getMessage());

    GridGrindResponse.SheetSchemaReport schema =
        new GridGrindResponse.SheetSchemaReport(
            "Budget",
            "A1",
            2,
            1,
            1,
            List.of(
                new GridGrindResponse.SchemaColumnReport(
                    0,
                    "A",
                    "Owner",
                    1,
                    0,
                    List.of(new GridGrindResponse.TypeCountReport("STRING", 1)),
                    "STRING")));
    assertEquals(1, schema.columns().size());
    assertEquals(
        "count must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.TypeCountReport("STRING", 0))
            .getMessage());

    GridGrindResponse.NamedRangeSurfaceReport surface =
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
                    GridGrindResponse.NamedRangeBackingKind.RANGE)));
    assertEquals(1, surface.namedRanges().size());
    assertEquals(
        "formulaCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.SheetFormulaSurfaceReport("Budget", -1, 0, List.of()))
            .getMessage());
    assertEquals(
        "formula must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.FormulaPatternReport(" ", 1, List.of("A1")))
            .getMessage());

    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(1, 0, 1, 0);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.NAMED_RANGE_UNRESOLVED_TARGET,
            AnalysisSeverity.INFO,
            "Formula-backed name",
            "Named range stores a formula.",
            new GridGrindResponse.AnalysisLocationReport.NamedRange(
                "BudgetTotal", new NamedRangeScope.Workbook()),
            List.of("Budget!$B$4"));
    assertEquals(
        1,
        new GridGrindResponse.WorkbookFindingsReport(summary, List.of(finding)).findings().size());
    assertEquals(
        "totalCount must equal errorCount + warningCount + infoCount",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.AnalysisSummaryReport(2, 0, 1, 0))
            .getMessage());
  }

  @Test
  void commentCellAndProblemReportsCoverDefaultsAndValidation() {
    GridGrindResponse.CommentReport comment =
        new GridGrindResponse.CommentReport(
            "Owner note",
            "Alice",
            true,
            List.of(new RichTextRunReport("Owner note", style().font())),
            new CommentAnchorReport(0, 0, 1, 2));
    assertEquals("Owner note", comment.text());
    assertEquals(
        "comment runs must concatenate to the plain text",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.CommentReport(
                        "Owner note",
                        "Alice",
                        false,
                        List.of(new RichTextRunReport("Mismatch", style().font())),
                        null))
            .getMessage());

    GridGrindResponse.CellReport numberCell =
        new GridGrindResponse.CellReport.NumberReport(
            "A1", "NUMERIC", "1", style(), null, null, 1.0d);
    GridGrindResponse.CellReport booleanCell =
        new GridGrindResponse.CellReport.BooleanReport(
            "A2", "BOOLEAN", "TRUE", style(), null, null, true);
    GridGrindResponse.CellReport errorCell =
        new GridGrindResponse.CellReport.ErrorReport(
            "A3", "ERROR", "#DIV/0!", style(), null, null, "#DIV/0!");
    GridGrindResponse.CellReport formulaCell =
        new GridGrindResponse.CellReport.FormulaReport(
            "A4",
            "FORMULA",
            "1",
            style(),
            null,
            null,
            "SUM(A1:A2)",
            new GridGrindResponse.CellReport.NumberReport(
                "A4", "NUMERIC", "1", style(), null, null, 1.0d));

    assertEquals(1.0d, numberCell.numberValue());
    assertEquals(true, booleanCell.booleanValue());
    assertEquals("#DIV/0!", errorCell.errorValue());
    assertEquals("SUM(A1:A2)", formulaCell.formula());
    assertNull(
        new GridGrindResponse.NamedRangeReport.FormulaReport(
                "Expr", new NamedRangeScope.Workbook(), "SUM(A1:A2)")
            .target());

    GridGrindResponse.Problem problem =
        GridGrindResponse.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE"));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST.category(), problem.category());
    assertEquals(
        "causes must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new GridGrindResponse.Problem(
                        GridGrindProblemCode.INVALID_REQUEST,
                        GridGrindProblemCode.INVALID_REQUEST.category(),
                        GridGrindProblemCode.INVALID_REQUEST.recovery(),
                        GridGrindProblemCode.INVALID_REQUEST.title(),
                        "bad request",
                        GridGrindProblemCode.INVALID_REQUEST.resolution(),
                        new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE"),
                        null,
                        java.util.Arrays.asList((GridGrindResponse.ProblemCause) null)))
            .getMessage());
  }

  private static GridGrindResponse.CellStyleReport style() {
    CellBorderSideReport emptySide = new CellBorderSideReport(ExcelBorderStyle.NONE, null);
    return new GridGrindResponse.CellStyleReport(
        "General",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            new FontHeightReport(220, BigDecimal.valueOf(11)),
            null,
            false,
            false),
        new CellFillReport(ExcelFillPattern.NONE, null, null),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }
}
