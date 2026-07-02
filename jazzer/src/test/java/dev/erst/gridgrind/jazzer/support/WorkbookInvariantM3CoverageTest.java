package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellReport;
import dev.erst.gridgrind.contract.dto.CellStyleReport;
import dev.erst.gridgrind.contract.dto.CellTemporalKind;
import dev.erst.gridgrind.contract.dto.CellTemporalReport;
import dev.erst.gridgrind.contract.dto.CellValueReport;
import dev.erst.gridgrind.contract.dto.CommentReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.RichTextRunReport;
import dev.erst.gridgrind.contract.dto.WindowDimensionsReport;
import dev.erst.gridgrind.contract.dto.WindowReport;
import dev.erst.gridgrind.contract.dto.WindowRowReport;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused invariant coverage for the projection-aware M3 cell and window contract. */
class WorkbookInvariantM3CoverageTest {
  @Test
  void acceptsSparseAndDenseWindowShapes() {
    assertDoesNotThrow(
        () ->
            WorkbookInvariantAnalysisLayoutChecks.requireWindowShape(
                new WindowReport.Sparse(
                    "Budget",
                    "A1",
                    new WindowDimensionsReport(2, 2),
                    List.of(textCell("A1", "Ada")))));
    assertDoesNotThrow(
        () ->
            WorkbookInvariantAnalysisLayoutChecks.requireWindowShape(
                new WindowReport.Dense(
                    "Budget",
                    "A1",
                    new WindowDimensionsReport(1, 2),
                    List.of(
                        new WindowRowReport(0, List.of(blankCell("A1"), textCell("B1", "Ada")))))));
  }

  @Test
  void acceptsProjectionAwareCellReportsAndFormulaEvaluations() {
    assertDoesNotThrow(
        () ->
            WorkbookInvariantCellContentChecks.requireCellReportShape(
                new CellReport.TextReport(
                    "A1",
                    Optional.of("Ada"),
                    Optional.of(style()),
                    Optional.of(new HyperlinkTarget.Url("https://example.com/report")),
                    Optional.of(new CommentReport("Review", "GridGrind", true)),
                    Optional.of("Ada"),
                    Optional.of(List.of(run("Ada"))))));
    assertDoesNotThrow(
        () ->
            WorkbookInvariantCellContentChecks.requireCellReportShape(
                new CellReport.NumberReport(
                    "A2",
                    Optional.of("7/1/2026"),
                    Optional.of(style()),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(46160.0d),
                    Optional.of(
                        CellTemporalReport.temporal(CellTemporalKind.DATE, "2026-07-01")))));
    assertDoesNotThrow(
        () ->
            WorkbookInvariantCellContentChecks.requireCellReportShape(
                new CellReport.FormulaReport(
                    "A3",
                    Optional.of("Ada"),
                    Optional.of(style()),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of("CONCAT(\"A\",\"da\")"),
                    Optional.of(
                        new CellValueReport.TextValue("Ada", Optional.of(List.of(run("Ada"))))))));
  }

  @Test
  void inspectionResultMatchingUsesWindowDimensions() {
    InspectionStep step =
        new InspectionStep(
            "window",
            new RangeSelector.RectangularWindow("Budget", "A1", 1, 1),
            new SheetIntrospectionQuery.GetWindow(Optional.empty(), true));
    SheetInspectionResult.WindowResult result =
        new SheetInspectionResult.WindowResult(
            "window",
            new WindowReport.Dense(
                "Budget",
                "A1",
                new WindowDimensionsReport(1, 1),
                List.of(new WindowRowReport(0, List.of(blankCell("A1"))))));

    assertDoesNotThrow(
        () -> WorkbookInvariantInspectionResultChecks.requireReadMatchesRequest(step, result));
  }

  private static CellReport.BlankReport blankCell(String address) {
    return new CellReport.BlankReport(
        address, Optional.empty(), Optional.empty(), Optional.empty(), Optional.empty());
  }

  private static CellReport.TextReport textCell(String address, String value) {
    return new CellReport.TextReport(
        address,
        Optional.of(value),
        Optional.of(style()),
        Optional.empty(),
        Optional.empty(),
        Optional.of(value),
        Optional.empty());
  }

  private static RichTextRunReport run(String text) {
    return new RichTextRunReport(text, style().font());
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
            "Calibri",
            new FontHeightReport(220, new BigDecimal("11")),
            null,
            false,
            false),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }
}
