package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Objects;

/** One reusable derived finding emitted by an analysis read. */
public record AnalysisFindingReport(
    AnalysisFindingCode code,
    AnalysisSeverity severity,
    String title,
    String message,
    AnalysisLocationReport location,
    List<String> evidence) {
  public AnalysisFindingReport {
    Objects.requireNonNull(code, "code must not be null");
    Objects.requireNonNull(severity, "severity must not be null");
    Objects.requireNonNull(title, "title must not be null");
    Objects.requireNonNull(message, "message must not be null");
    Objects.requireNonNull(location, "location must not be null");
    if (title.isBlank()) {
      throw new IllegalArgumentException("title must not be blank");
    }
    if (message.isBlank()) {
      throw new IllegalArgumentException("message must not be blank");
    }
    evidence = WorkbookResultSupport.copyStrings(evidence, "evidence");
  }
}
