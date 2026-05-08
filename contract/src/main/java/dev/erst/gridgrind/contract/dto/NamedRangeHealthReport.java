package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Named-range-health analysis for one selected named-range set. */
public record NamedRangeHealthReport(
    int checkedNamedRangeCount,
    AnalysisSummaryReport summary,
    List<AnalysisFindingReport> findings) {
  public NamedRangeHealthReport {
    if (checkedNamedRangeCount < 0) {
      throw new IllegalArgumentException("checkedNamedRangeCount must not be negative");
    }
    Objects.requireNonNull(summary, "summary must not be null");
    findings = GridGrindResponseSupport.copyValues(findings, "findings");
  }
}
