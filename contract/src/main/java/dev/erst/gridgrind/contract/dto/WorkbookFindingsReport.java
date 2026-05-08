package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Aggregated workbook findings composed from all shipped analysis families. */
public record WorkbookFindingsReport(
    AnalysisSummaryReport summary, List<AnalysisFindingReport> findings) {
  public WorkbookFindingsReport {
    Objects.requireNonNull(summary, "summary must not be null");
    findings = GridGrindResponseSupport.copyValues(findings, "findings");
  }
}
