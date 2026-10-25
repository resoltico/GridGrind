package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookFindingsReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for step-keyed inspection result payloads. */
class InspectionResultTest {
  @Test
  void sheetsResultCopiesSheetNamesAndRejectsBlankStepId() {
    List<SheetLabel> names =
        new ArrayList<>(List.of(new SheetLabel("Budget"), new SheetLabel("Report")));
    WorkbookInspectionResult.SheetsResult result =
        new WorkbookInspectionResult.SheetsResult(new StepId("sheets"), names);
    names.clear();

    assertEquals("sheets", result.stepId().value());
    assertEquals(
        List.of("Budget", "Report"),
        result.sheetNames().stream()
            .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
            .toList());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookInspectionResult.SheetsResult(new StepId(" "), List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookInspectionResult.SheetsResult(new StepId("sheets"), null));
  }

  @Test
  void introspectionResultsUseStepIdsAndCopyCollections() {
    List<dev.erst.gridgrind.core.model.value.CellReport> cells = new ArrayList<>();
    cells.add(
        new dev.erst.gridgrind.core.model.value.CellReport.BlankReport(
            "A1",
            java.util.Optional.of(""),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.empty(),
            java.util.Optional.empty()));
    SheetInspectionResult.CellsResult result =
        new SheetInspectionResult.CellsResult(new StepId("cells"), new SheetLabel("Budget"), cells);

    cells.clear();

    assertEquals("cells", result.stepId().value());
    assertEquals("A1", result.cells().getFirst().address());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookInspectionResult.WorkbookProtectionResult(
                new StepId(" "), new WorkbookProtectionReport(false, false, false, false, false)));
  }

  @Test
  void analysisResultsRetainReportsAndRejectBlankStepIds() {
    AnalysisSummaryReport summary = new AnalysisSummaryReport(1, 0, 0, 1);
    AnalysisFindingReport finding =
        new AnalysisFindingReport(
            AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            AnalysisSeverity.INFO,
            "Volatile formula",
            "Formula uses NOW().",
            new AnalysisLocationReport.Workbook(),
            List.of("NOW()"));
    WorkbookAnalysisResult.WorkbookFindingsResult findings =
        new WorkbookAnalysisResult.WorkbookFindingsResult(
            new StepId("workbook-findings"), new WorkbookFindingsReport(summary, List.of(finding)));

    assertEquals(1, findings.analysis().summary().totalCount());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookAnalysisResult.WorkbookFindingsResult(
                new StepId(" "), findings.analysis()));
    assertThrows(
        NullPointerException.class,
        () ->
            new WorkbookAnalysisResult.HyperlinkHealthResult(new StepId("hyperlink-health"), null));
  }

  private static CellStyleReport minimalStyle() {
    CellBorderSideReport emptySide = new CellBorderSideReport.None();
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
