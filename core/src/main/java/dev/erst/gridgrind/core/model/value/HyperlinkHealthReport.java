package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookResultSupport;
import java.util.List;
import java.util.Objects;

/** Hyperlink-health analysis for one selected sheet set. */
public record HyperlinkHealthReport(
    int checkedHyperlinkCount,
    AnalysisSummaryReport summary,
    List<AnalysisFindingReport> findings) {
  public HyperlinkHealthReport {
    if (checkedHyperlinkCount < 0) {
      throw new IllegalArgumentException("checkedHyperlinkCount must not be negative");
    }
    Objects.requireNonNull(summary, "summary must not be null");
    findings = WorkbookResultSupport.copyValues(findings, "findings");
  }
}
