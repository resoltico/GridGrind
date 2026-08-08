package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Coverage tests for nested response-report records that are not exercised by end-to-end flows. */
class WorkbookResultNestedCoverageTest {
  @Test
  void windowLayoutSchemaAndAnalysisReportsValidateAndCopyCollections() {
    dev.erst.gridgrind.contract.dto.CellReport textCell =
        new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
            "A1",
            Optional.of("Owner"),
            Optional.of(style()),
            Optional.empty(),
            Optional.empty(),
            Optional.of("Owner"),
            Optional.of(List.of(new RichTextRunReport("Owner", style().font()))));
    WindowRowReport row = new WindowRowReport(0, List.of(textCell));
    WindowReport window =
        new WindowReport.Dense("Budget", "A1", new WindowDimensionsReport(1, 1), List.of(row));
    assertEquals("Budget", window.sheetName());
    assertEquals(
        "rowCount must be greater than 0",
        assertThrows(IllegalArgumentException.class, () -> new WindowDimensionsReport(0, 1))
            .getMessage());

    SheetLayoutReport layout =
        new SheetLayoutReport(
            "Budget",
            new PaneReport.Split(1, 1, 0, 0, ExcelPaneRegion.LOWER_RIGHT),
            125,
            SheetPresentationReport.defaults(),
            List.of(new ColumnLayoutReport(0, 8.43d, false, 0, false)),
            List.of(new RowLayoutReport(0, 15.0d, false, 0, false)));
    assertEquals(125, layout.zoomPercent());
    assertEquals(
        "zoomPercent must be between 10 and 400 inclusive: 401",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new SheetLayoutReport(
                        "Budget",
                        new PaneReport.None(),
                        401,
                        SheetPresentationReport.defaults(),
                        List.of(),
                        List.of()))
            .getMessage());

    SheetSchemaReport schema =
        new SheetSchemaReport(
            "Budget",
            "A1",
            2,
            1,
            1,
            List.of(
                new SchemaColumnReport(
                    0, "A", "Owner", 1, 0, List.of(new TypeCountReport("TEXT", 1)), "TEXT")));
    assertEquals(1, schema.columns().size());
    SchemaColumnReport optionalDominantTypeColumn =
        new SchemaColumnReport(
            1, "B", "Cost", 1, 0, List.of(new TypeCountReport("NUMBER", 1)), null);
    assertNull(optionalDominantTypeColumn.dominantType());
    assertEquals(
        "count must be greater than 0",
        assertThrows(IllegalArgumentException.class, () -> new TypeCountReport("TEXT", 0))
            .getMessage());
    assertEquals(
        "type must be one of TEXT, NUMBER, BOOLEAN, ERROR, DATE, TIME, DATE_TIME but was STRING",
        assertThrows(IllegalArgumentException.class, () -> new TypeCountReport("STRING", 1))
            .getMessage());

    NamedRangeSurfaceReport surface =
        new NamedRangeSurfaceReport(
            1,
            0,
            1,
            0,
            List.of(
                new NamedRangeSurfaceEntryReport(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    "Budget!$B$4",
                    NamedRangeBackingKind.RANGE)));
    assertEquals(1, surface.namedRanges().size());
    assertEquals(
        "formulaCellCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SheetFormulaSurfaceReport("Budget", -1, 0, List.of()))
            .getMessage());
    assertEquals(
        "formula must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new FormulaPatternReport(" ", 1, List.of("A1")))
            .getMessage());

    AnalysisSummaryReport summary = new AnalysisSummaryReport(1, 0, 1, 0);
    AnalysisFindingReport finding =
        new AnalysisFindingReport(
            AnalysisFindingCode.NAMED_RANGE_UNRESOLVED_TARGET,
            AnalysisSeverity.INFO,
            "Formula-backed name",
            "Named range stores a formula.",
            new AnalysisLocationReport.NamedRange("BudgetTotal", new NamedRangeScope.Workbook()),
            List.of("Budget!$B$4"));
    assertEquals(1, new WorkbookFindingsReport(summary, List.of(finding)).findings().size());
    assertEquals(
        "totalCount must equal errorCount + warningCount + infoCount",
        assertThrows(IllegalArgumentException.class, () -> new AnalysisSummaryReport(2, 0, 1, 0))
            .getMessage());
  }

  @Test
  void commentCellAndProblemReportsCoverDefaultsAndValidation() {
    CommentReport comment =
        new CommentReport(
            "Owner note",
            "Alice",
            true,
            Optional.of(List.of(new RichTextRunReport("Owner note", style().font()))),
            Optional.of(new CommentAnchorReport(0, 0, 1, 2)));
    assertEquals("Owner note", comment.text());
    assertEquals(
        "comment runs must concatenate to the plain text",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CommentReport(
                        "Owner note",
                        "Alice",
                        false,
                        Optional.of(List.of(new RichTextRunReport("Mismatch", style().font()))),
                        Optional.empty()))
            .getMessage());

    dev.erst.gridgrind.contract.dto.CellReport.NumberReport numberCell =
        new dev.erst.gridgrind.contract.dto.CellReport.NumberReport(
            "A1",
            java.util.Optional.of("1"),
            java.util.Optional.of(style()),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.of(1.0d),
            java.util.Optional.empty());
    dev.erst.gridgrind.contract.dto.CellReport.BooleanReport booleanCell =
        new dev.erst.gridgrind.contract.dto.CellReport.BooleanReport(
            "A2",
            java.util.Optional.of("TRUE"),
            java.util.Optional.of(style()),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.of(true));
    dev.erst.gridgrind.contract.dto.CellReport.ErrorReport errorCell =
        new dev.erst.gridgrind.contract.dto.CellReport.ErrorReport(
            "A3",
            java.util.Optional.of("#DIV/0!"),
            java.util.Optional.of(style()),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.of("#DIV/0!"));
    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formulaCell =
        new dev.erst.gridgrind.contract.dto.CellReport.FormulaReport(
            "A4",
            java.util.Optional.of("1"),
            java.util.Optional.of(style()),
            java.util.Optional.empty(),
            java.util.Optional.empty(),
            java.util.Optional.of("SUM(A1:A2)"),
            java.util.Optional.of(
                new CellValueReport.NumberValue(1.0d, java.util.Optional.empty())));

    assertEquals(1.0d, numberCell.numberValue().orElseThrow());
    assertEquals(true, booleanCell.booleanValue().orElseThrow());
    assertEquals("#DIV/0!", errorCell.errorValue().orElseThrow());
    assertEquals("SUM(A1:A2)", formulaCell.formula().orElseThrow());
    assertInstanceOf(
        NamedRangeReport.FormulaReport.class,
        new NamedRangeReport.FormulaReport("Expr", new NamedRangeScope.Workbook(), "SUM(A1:A2)"));

    GridGrindProblemDetail.Problem problem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape.known(
                    "NEW", "NONE")));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST.category(), problem.category());
    assertEquals(
        "causes must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new GridGrindProblemDetail.Problem(
                        GridGrindProblemCode.INVALID_REQUEST,
                        GridGrindProblemCode.INVALID_REQUEST.category(),
                        GridGrindProblemCode.INVALID_REQUEST.recovery(),
                        GridGrindProblemCode.INVALID_REQUEST.title(),
                        "bad request",
                        GridGrindProblemCode.INVALID_REQUEST.resolution(),
                        new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces
                                .RequestShape.known("NEW", "NONE")),
                        java.util.Optional.empty(),
                        java.util.Arrays.asList((GridGrindProblemDetail.ProblemCause) null)))
            .getMessage());
  }

  private static CellStyleReport style() {
    CellBorderSideReport emptySide = new CellBorderSideReport(ExcelBorderStyle.NONE, null);
    return new CellStyleReport(
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
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }
}
