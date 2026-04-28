package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
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
            AnalysisFindingCode.TABLE_OVERLAPPING_RANGE,
            AnalysisSeverity.ERROR,
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
