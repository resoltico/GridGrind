package dev.erst.gridgrind.contract.dto;

import java.util.List;
import java.util.Objects;

/** Formula-health analysis for one selected sheet set. */
public record FormulaHealthReport(
    int checkedFormulaCellCount,
    AnalysisSummaryReport summary,
    List<AnalysisFindingReport> findings) {
  public FormulaHealthReport {
    if (checkedFormulaCellCount < 0) {
      throw new IllegalArgumentException("checkedFormulaCellCount must not be negative");
    }
    Objects.requireNonNull(summary, "summary must not be null");
    findings = GridGrindResponseSupport.copyValues(findings, "findings");
  }
}
