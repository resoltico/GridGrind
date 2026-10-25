package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.assertion.AnalysisAssertion;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.assertion.CompositeAssertion;
import dev.erst.gridgrind.contract.assertion.PresenceAssertion;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Canonical assertion helpers kept internal to the focused Java authoring surface. */
final class Checks {
  private Checks() {}

  static PresenceAssertion.NamedRangePresent namedRangePresent() {
    return new PresenceAssertion.NamedRangePresent();
  }

  static PresenceAssertion.NamedRangeAbsent namedRangeAbsent() {
    return new PresenceAssertion.NamedRangeAbsent();
  }

  static PresenceAssertion.TablePresent tablePresent() {
    return new PresenceAssertion.TablePresent();
  }

  static PresenceAssertion.TableAbsent tableAbsent() {
    return new PresenceAssertion.TableAbsent();
  }

  static PresenceAssertion.PivotTablePresent pivotTablePresent() {
    return new PresenceAssertion.PivotTablePresent();
  }

  static PresenceAssertion.PivotTableAbsent pivotTableAbsent() {
    return new PresenceAssertion.PivotTableAbsent();
  }

  static PresenceAssertion.ChartPresent chartPresent() {
    return new PresenceAssertion.ChartPresent();
  }

  static PresenceAssertion.ChartAbsent chartAbsent() {
    return new PresenceAssertion.ChartAbsent();
  }

  static CellAssertion.CellValue cellValue(ExpectedValues.Value expectedValue) {
    return new CellAssertion.CellValue(ExpectedValues.toCellScalarValue(expectedValue));
  }

  static CellAssertion.DisplayValue displayValue(String displayValue) {
    return new CellAssertion.DisplayValue(displayValue);
  }

  static CellAssertion.FormulaText formulaText(String formula) {
    return new CellAssertion.FormulaText(formula);
  }

  static AnalysisAssertion.AnalysisMaxSeverity analysisMaxSeverity(
      InspectionQuery.Analysis query, AnalysisSeverity maximumSeverity) {
    return new AnalysisAssertion.AnalysisMaxSeverity(query, maximumSeverity);
  }

  static AnalysisAssertion.AnalysisFindingPresent analysisFindingPresent(
      InspectionQuery.Analysis query,
      AnalysisFindingCode code,
      AnalysisSeverity severity,
      String messageContains) {
    return new AnalysisAssertion.AnalysisFindingPresent(
        query, code, Optional.ofNullable(severity), Optional.ofNullable(messageContains));
  }

  static AnalysisAssertion.AnalysisFindingAbsent analysisFindingAbsent(
      InspectionQuery.Analysis query,
      AnalysisFindingCode code,
      AnalysisSeverity severity,
      String messageContains) {
    return new AnalysisAssertion.AnalysisFindingAbsent(
        query, code, Optional.ofNullable(severity), Optional.ofNullable(messageContains));
  }

  static CompositeAssertion.AllOf allOf(Assertion... assertions) {
    return new CompositeAssertion.AllOf(List.of(assertions));
  }

  static CompositeAssertion.AnyOf anyOf(Assertion... assertions) {
    return new CompositeAssertion.AnyOf(List.of(assertions));
  }

  static CompositeAssertion.Not not(Assertion assertion) {
    return new CompositeAssertion.Not(
        Objects.requireNonNull(assertion, "assertion must not be null"));
  }
}
