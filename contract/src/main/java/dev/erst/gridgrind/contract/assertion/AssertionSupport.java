package dev.erst.gridgrind.contract.assertion;

import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Shared validation helpers for assertion records and failure payloads. */
final class AssertionSupport {
  private AssertionSupport() {}

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static List<Assertion> copyAssertions(List<Assertion> assertions, String fieldName) {
    Objects.requireNonNull(assertions, fieldName + " must not be null");
    List<Assertion> copy = new ArrayList<>(assertions.size());
    for (Assertion assertion : assertions) {
      copy.add(Objects.requireNonNull(assertion, fieldName + " must not contain nulls"));
    }
    if (copy.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return List.copyOf(copy);
  }

  static List<InspectionResult> copyObservations(
      List<InspectionResult> observations, String fieldName) {
    Objects.requireNonNull(observations, fieldName + " must not be null");
    List<InspectionResult> copy = new ArrayList<>(observations.size());
    for (InspectionResult observation : observations) {
      copy.add(Objects.requireNonNull(observation, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  static List<GridGrindWorkbookSurfaceReports.NamedRangeReport> copyNamedRanges(
      List<GridGrindWorkbookSurfaceReports.NamedRangeReport> namedRanges, String fieldName) {
    return copyValues(namedRanges, fieldName);
  }

  static List<TableEntryReport> copyTables(List<TableEntryReport> tables, String fieldName) {
    return copyValues(tables, fieldName);
  }

  static List<PivotTableReport> copyPivotTables(
      List<PivotTableReport> pivotTables, String fieldName) {
    return copyValues(pivotTables, fieldName);
  }

  static List<ChartReport> copyCharts(List<ChartReport> charts, String fieldName) {
    return copyValues(charts, fieldName);
  }

  static InspectionQuery.Analysis requireAnalysisQuery(InspectionQuery query, String fieldName) {
    Objects.requireNonNull(query, fieldName + " must not be null");
    if (!(query instanceof InspectionQuery.Analysis analysis)) {
      throw new IllegalArgumentException(fieldName + " must be an analysis query");
    }
    return analysis;
  }

  static AnalysisFindingCode requireFindingCode(AnalysisFindingCode code, String fieldName) {
    return Objects.requireNonNull(code, fieldName + " must not be null");
  }

  static AnalysisSeverity requireSeverity(AnalysisSeverity severity, String fieldName) {
    return Objects.requireNonNull(severity, fieldName + " must not be null");
  }

  static Double requireFiniteNumber(Double number, String fieldName) {
    Objects.requireNonNull(number, fieldName + " must not be null");
    if (!Double.isFinite(number)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    return number;
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
