package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookResultSupport;
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
    findings = WorkbookResultSupport.copyValues(findings, "findings");
  }
}
