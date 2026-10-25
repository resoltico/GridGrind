package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.layout.WindowDimensionsReport;
import dev.erst.gridgrind.core.model.layout.WindowReport;
import dev.erst.gridgrind.core.model.layout.WindowRowReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellTemporalKind;
import dev.erst.gridgrind.core.model.value.CellTemporalReport;
import dev.erst.gridgrind.core.model.value.CellValueReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
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
                    new SheetLabel("Budget"),
                    new CellAddress("A1"),
                    new WindowDimensionsReport(2, 2),
                    List.of(textCell("A1", "Ada")))));
    assertDoesNotThrow(
        () ->
            WorkbookInvariantAnalysisLayoutChecks.requireWindowShape(
                new WindowReport.Dense(
                    new SheetLabel("Budget"),
                    new CellAddress("A1"),
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
                    new CellAddress("A1"),
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
                    new CellAddress("A2"),
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
                    new CellAddress("A3"),
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
            new StepId("window"),
            new RangeSelector.RectangularWindow(
                new SheetName("Budget"), new CellAddress("A1"), 1, 1),
            new SheetIntrospectionQuery.GetWindow(Optional.empty(), true));
    SheetInspectionResult.WindowResult result =
        new SheetInspectionResult.WindowResult(
            new StepId("window"),
            new WindowReport.Dense(
                new SheetLabel("Budget"),
                new CellAddress("A1"),
                new WindowDimensionsReport(1, 1),
                List.of(new WindowRowReport(0, List.of(blankCell("A1"))))));

    assertDoesNotThrow(
        () -> WorkbookInvariantInspectionResultChecks.requireReadMatchesRequest(step, result));
  }

  private static CellReport.BlankReport blankCell(String address) {
    return new CellReport.BlankReport(
        new CellAddress(address),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        Optional.empty());
  }

  private static CellReport.TextReport textCell(String address, String value) {
    return new CellReport.TextReport(
        new CellAddress(address),
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
    CellBorderSideReport emptySide = new CellBorderSideReport.None();
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
