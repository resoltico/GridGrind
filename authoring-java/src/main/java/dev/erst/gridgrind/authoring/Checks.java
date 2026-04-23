package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.util.List;
import java.util.Objects;

/** Canonical assertion helpers kept internal to the focused Java authoring surface. */
final class Checks {
  private Checks() {}

  static Assertion.Present present() {
    return new Assertion.Present();
  }

  static Assertion.Absent absent() {
    return new Assertion.Absent();
  }

  static Assertion.CellValue cellValue(Values.ExpectedValue expectedValue) {
    return new Assertion.CellValue(Values.toExpectedCellValue(expectedValue));
  }

  static Assertion.DisplayValue displayValue(String displayValue) {
    return new Assertion.DisplayValue(displayValue);
  }

  static Assertion.FormulaText formulaText(String formula) {
    return new Assertion.FormulaText(formula);
  }

  static Assertion.AnalysisMaxSeverity analysisMaxSeverity(
      InspectionQuery.Analysis query, AnalysisSeverity maximumSeverity) {
    return new Assertion.AnalysisMaxSeverity(query, maximumSeverity);
  }

  static Assertion.AnalysisFindingPresent analysisFindingPresent(
      InspectionQuery.Analysis query,
      AnalysisFindingCode code,
      AnalysisSeverity severity,
      String messageContains) {
    return new Assertion.AnalysisFindingPresent(query, code, severity, messageContains);
  }

  static Assertion.AnalysisFindingAbsent analysisFindingAbsent(
      InspectionQuery.Analysis query,
      AnalysisFindingCode code,
      AnalysisSeverity severity,
      String messageContains) {
    return new Assertion.AnalysisFindingAbsent(query, code, severity, messageContains);
  }

  static Assertion.AllOf allOf(Assertion... assertions) {
    return new Assertion.AllOf(List.of(assertions));
  }

  static Assertion.AnyOf anyOf(Assertion... assertions) {
    return new Assertion.AnyOf(List.of(assertions));
  }

  static Assertion.Not not(Assertion assertion) {
    return new Assertion.Not(Objects.requireNonNull(assertion, "assertion must not be null"));
  }
}
