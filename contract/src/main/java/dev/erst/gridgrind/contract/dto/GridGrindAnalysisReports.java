package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Objects;

/** Derived analysis and health report families. */
public interface GridGrindAnalysisReports {
  /** Summary counts for one finding-bearing analysis result. */
  record AnalysisSummaryReport(int totalCount, int errorCount, int warningCount, int infoCount) {
    public AnalysisSummaryReport {
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

  /** Precise workbook location attached to one derived analysis finding. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = AnalysisLocationReport.Workbook.class, name = "WORKBOOK"),
    @JsonSubTypes.Type(value = AnalysisLocationReport.Sheet.class, name = "SHEET"),
    @JsonSubTypes.Type(value = AnalysisLocationReport.Cell.class, name = "CELL"),
    @JsonSubTypes.Type(value = AnalysisLocationReport.Range.class, name = "RANGE"),
    @JsonSubTypes.Type(value = AnalysisLocationReport.NamedRange.class, name = "NAMED_RANGE")
  })
  sealed interface AnalysisLocationReport
      permits AnalysisLocationReport.Workbook,
          AnalysisLocationReport.Sheet,
          AnalysisLocationReport.Cell,
          AnalysisLocationReport.Range,
          AnalysisLocationReport.NamedRange {

    /** Workbook-level finding with no narrower location. */
    record Workbook() implements AnalysisLocationReport {}

    /** One whole-sheet finding. */
    record Sheet(String sheetName) implements AnalysisLocationReport {
      public Sheet {
        Objects.requireNonNull(sheetName, "sheetName must not be null");
        if (sheetName.isBlank()) {
          throw new IllegalArgumentException("sheetName must not be blank");
        }
      }
    }

    /** One concrete cell finding. */
    record Cell(String sheetName, String address) implements AnalysisLocationReport {
      public Cell {
        Objects.requireNonNull(sheetName, "sheetName must not be null");
        Objects.requireNonNull(address, "address must not be null");
        if (sheetName.isBlank()) {
          throw new IllegalArgumentException("sheetName must not be blank");
        }
        if (address.isBlank()) {
          throw new IllegalArgumentException("address must not be blank");
        }
      }
    }

    /** One rectangular range finding. */
    record Range(String sheetName, String range) implements AnalysisLocationReport {
      public Range {
        Objects.requireNonNull(sheetName, "sheetName must not be null");
        Objects.requireNonNull(range, "range must not be null");
        if (sheetName.isBlank()) {
          throw new IllegalArgumentException("sheetName must not be blank");
        }
        if (range.isBlank()) {
          throw new IllegalArgumentException("range must not be blank");
        }
      }
    }

    /** One named-range finding. */
    record NamedRange(String name, NamedRangeScope scope) implements AnalysisLocationReport {
      public NamedRange {
        Objects.requireNonNull(name, "name must not be null");
        Objects.requireNonNull(scope, "scope must not be null");
        if (name.isBlank()) {
          throw new IllegalArgumentException("name must not be blank");
        }
      }
    }
  }

  /** One reusable derived finding emitted by an analysis read. */
  record AnalysisFindingReport(
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
      evidence = GridGrindResponseSupport.copyStrings(evidence, "evidence");
    }
  }

  /** Formula-health analysis for one selected sheet set. */
  record FormulaHealthReport(
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

  /** Hyperlink-health analysis for one selected sheet set. */
  record HyperlinkHealthReport(
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

  /** Named-range-health analysis for one selected named-range set. */
  record NamedRangeHealthReport(
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

  /** Aggregated workbook findings composed from all shipped analysis families. */
  record WorkbookFindingsReport(
      AnalysisSummaryReport summary, List<AnalysisFindingReport> findings) {
    public WorkbookFindingsReport {
      Objects.requireNonNull(summary, "summary must not be null");
      findings = GridGrindResponseSupport.copyValues(findings, "findings");
    }
  }
}
