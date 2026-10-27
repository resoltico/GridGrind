package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
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
    TextValidation.requireNonBlank(title, "title");
    TextValidation.requireNonBlank(message, "message");
    evidence = WorkbookResultSupport.copyStrings(evidence, "evidence");
  }
}
