package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.result.WorkbookResultSupport;
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
    findings = WorkbookResultSupport.copyValues(findings, "findings");
  }
}
