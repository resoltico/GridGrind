package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for table-health protocol reports. */
class TableHealthReportTest {
  @Test
  void validatesCountsAndCopiesFindings() {
    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(1, 0, 1, 0);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode.TABLE_OVERLAPPING_RANGE,
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.ERROR,
            "Table range overlaps another table",
            "Table metadata overlaps one or more other tables on the same sheet.",
            new GridGrindResponse.AnalysisLocationReport.Range("Budget", "A1:C4"),
            List.of("BudgetTable@A1:C4", "Queue@B1:D4"));

    TableHealthReport report = new TableHealthReport(1, summary, List.of(finding));

    assertEquals(List.of(finding), report.findings());
    assertThrows(
        IllegalArgumentException.class, () -> new TableHealthReport(-1, summary, List.of(finding)));
    assertThrows(
        NullPointerException.class, () -> new TableHealthReport(1, null, List.of(finding)));
    assertThrows(
        NullPointerException.class,
        () -> new TableHealthReport(1, summary, List.of(finding, null)));
  }
}
