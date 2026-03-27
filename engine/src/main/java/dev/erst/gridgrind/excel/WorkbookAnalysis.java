package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Shared engine-side analysis payload model reused across finding-bearing read results. */
public sealed interface WorkbookAnalysis
    permits WorkbookAnalysis.FormulaHealth,
        WorkbookAnalysis.HyperlinkHealth,
        WorkbookAnalysis.NamedRangeHealth,
        WorkbookAnalysis.WorkbookFindings {

  /** Severity of one derived workbook finding. */
  enum AnalysisSeverity {
    INFO,
    WARNING,
    ERROR
  }

  /** Stable machine-readable code for one derived workbook finding. */
  enum AnalysisFindingCode {
    FORMULA_ERROR_RESULT,
    FORMULA_EVALUATION_FAILURE,
    FORMULA_EXTERNAL_REFERENCE,
    FORMULA_VOLATILE_FUNCTION,
    HYPERLINK_MALFORMED_TARGET,
    HYPERLINK_MISSING_DOCUMENT_SHEET,
    HYPERLINK_INVALID_DOCUMENT_TARGET,
    NAMED_RANGE_BROKEN_REFERENCE,
    NAMED_RANGE_UNRESOLVED_TARGET,
    NAMED_RANGE_SCOPE_SHADOWING
  }

  /** Precise workbook location attached to one derived analysis finding. */
  sealed interface AnalysisLocation
      permits AnalysisLocation.Workbook,
          AnalysisLocation.Sheet,
          AnalysisLocation.Cell,
          AnalysisLocation.Range,
          AnalysisLocation.NamedRange {

    /** Workbook-level finding with no narrower location. */
    record Workbook() implements AnalysisLocation {}

    /** One whole-sheet finding. */
    record Sheet(String sheetName) implements AnalysisLocation {
      public Sheet {
        sheetName = requireNonBlank(sheetName, "sheetName");
      }
    }

    /** One concrete cell finding. */
    record Cell(String sheetName, String address) implements AnalysisLocation {
      public Cell {
        sheetName = requireNonBlank(sheetName, "sheetName");
        address = requireNonBlank(address, "address");
      }
    }

    /** One rectangular range finding. */
    record Range(String sheetName, String range) implements AnalysisLocation {
      public Range {
        sheetName = requireNonBlank(sheetName, "sheetName");
        range = requireNonBlank(range, "range");
      }
    }

    /** One named-range finding. */
    record NamedRange(String name, ExcelNamedRangeScope scope) implements AnalysisLocation {
      public NamedRange {
        name = requireNonBlank(name, "name");
        Objects.requireNonNull(scope, "scope must not be null");
      }
    }
  }

  /** One reusable derived finding emitted by an analysis read. */
  record AnalysisFinding(
      AnalysisFindingCode code,
      AnalysisSeverity severity,
      String title,
      String message,
      AnalysisLocation location,
      List<String> evidence) {
    public AnalysisFinding {
      Objects.requireNonNull(code, "code must not be null");
      Objects.requireNonNull(severity, "severity must not be null");
      title = requireNonBlank(title, "title");
      message = requireNonBlank(message, "message");
      Objects.requireNonNull(location, "location must not be null");
      evidence = copyStrings(evidence, "evidence");
    }
  }

  /** Summary counts for one finding-bearing analysis result. */
  record AnalysisSummary(int totalCount, int errorCount, int warningCount, int infoCount) {
    public AnalysisSummary {
      if (totalCount < 0) {
        throw new IllegalArgumentException("totalCount must not be negative");
      }
      if (errorCount < 0) {
        throw new IllegalArgumentException("errorCount must not be negative");
      }
      if (warningCount < 0) {
        throw new IllegalArgumentException("warningCount must not be negative");
      }
      if (infoCount < 0) {
        throw new IllegalArgumentException("infoCount must not be negative");
      }
      if (totalCount != errorCount + warningCount + infoCount) {
        throw new IllegalArgumentException(
            "totalCount must equal errorCount + warningCount + infoCount");
      }
    }
  }

  /** Formula-health analysis for one selected sheet set. */
  record FormulaHealth(
      int checkedFormulaCellCount, AnalysisSummary summary, List<AnalysisFinding> findings)
      implements WorkbookAnalysis {
    public FormulaHealth {
      if (checkedFormulaCellCount < 0) {
        throw new IllegalArgumentException("checkedFormulaCellCount must not be negative");
      }
      Objects.requireNonNull(summary, "summary must not be null");
      findings = copyValues(findings, "findings");
    }
  }

  /** Hyperlink-health analysis for one selected sheet set. */
  record HyperlinkHealth(
      int checkedHyperlinkCount, AnalysisSummary summary, List<AnalysisFinding> findings)
      implements WorkbookAnalysis {
    public HyperlinkHealth {
      if (checkedHyperlinkCount < 0) {
        throw new IllegalArgumentException("checkedHyperlinkCount must not be negative");
      }
      Objects.requireNonNull(summary, "summary must not be null");
      findings = copyValues(findings, "findings");
    }
  }

  /** Named-range-health analysis for one selected named-range set. */
  record NamedRangeHealth(
      int checkedNamedRangeCount, AnalysisSummary summary, List<AnalysisFinding> findings)
      implements WorkbookAnalysis {
    public NamedRangeHealth {
      if (checkedNamedRangeCount < 0) {
        throw new IllegalArgumentException("checkedNamedRangeCount must not be negative");
      }
      Objects.requireNonNull(summary, "summary must not be null");
      findings = copyValues(findings, "findings");
    }
  }

  /** Aggregated workbook findings composed from the first analysis family. */
  record WorkbookFindings(AnalysisSummary summary, List<AnalysisFinding> findings)
      implements WorkbookAnalysis {
    public WorkbookFindings {
      Objects.requireNonNull(summary, "summary must not be null");
      findings = copyValues(findings, "findings");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      Objects.requireNonNull(value, fieldName + " must not contain nulls");
    }
    return copy;
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = List.copyOf(values);
    for (T value : copy) {
      Objects.requireNonNull(value, fieldName + " must not contain nulls");
    }
    return copy;
  }
}
