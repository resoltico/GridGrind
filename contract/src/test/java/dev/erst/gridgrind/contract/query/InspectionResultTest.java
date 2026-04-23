package dev.erst.gridgrind.contract.query;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for step-keyed inspection result payloads. */
class InspectionResultTest {
  @Test
  void introspectionResultsUseStepIdsAndCopyCollections() {
    List<GridGrindResponse.CellReport> cells = new ArrayList<>();
    cells.add(
        new GridGrindResponse.CellReport.BlankReport(
            "A1", "BLANK", "", minimalStyle(), null, null));
    InspectionResult.CellsResult result =
        new InspectionResult.CellsResult("cells", "Budget", cells);

    cells.clear();

    assertEquals("cells", result.stepId());
    assertEquals("A1", result.cells().getFirst().address());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new InspectionResult.WorkbookProtectionResult(
                " ", new WorkbookProtectionReport(false, false, false, false, false)));
  }

  @Test
  void analysisResultsRetainReportsAndRejectBlankStepIds() {
    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(1, 0, 0, 1);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            AnalysisSeverity.INFO,
            "Volatile formula",
            "Formula uses NOW().",
            new GridGrindResponse.AnalysisLocationReport.Workbook(),
            List.of("NOW()"));
    InspectionResult.WorkbookFindingsResult findings =
        new InspectionResult.WorkbookFindingsResult(
            "workbook-findings",
            new GridGrindResponse.WorkbookFindingsReport(summary, List.of(finding)));

    assertEquals(1, findings.analysis().summary().totalCount());
    assertThrows(
        IllegalArgumentException.class,
        () -> new InspectionResult.WorkbookFindingsResult(" ", findings.analysis()));
    assertThrows(
        NullPointerException.class,
        () -> new InspectionResult.HyperlinkHealthResult("hyperlink-health", null));
  }

  private static GridGrindResponse.CellStyleReport minimalStyle() {
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
