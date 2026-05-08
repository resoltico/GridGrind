package dev.erst.gridgrind.contract.dto;

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
    findings = GridGrindResponseSupport.copyValues(findings, "findings");
  }
}
