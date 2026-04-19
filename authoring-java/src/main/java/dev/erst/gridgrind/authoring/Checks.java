package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.util.List;

/** Canonical assertion builders for the Java authoring layer. */
public final class Checks {
  private Checks() {}

  /** Returns one presence assertion for the targeted selector. */
  public static Assertion.Present present() {
    return new Assertion.Present();
  }

  /** Returns one absence assertion for the targeted selector. */
  public static Assertion.Absent absent() {
    return new Assertion.Absent();
  }

  /** Returns one effective cell-value assertion. */
  public static Assertion.CellValue cellValue(ExpectedCellValue expectedValue) {
    return new Assertion.CellValue(expectedValue);
  }

  /** Returns one rendered display-value assertion. */
  public static Assertion.DisplayValue displayValue(String displayValue) {
    return new Assertion.DisplayValue(displayValue);
  }

  /** Returns one formula-text assertion. */
  public static Assertion.FormulaText formulaText(String formula) {
    return new Assertion.FormulaText(formula);
  }

  /** Returns one full cell-style assertion. */
  public static Assertion.CellStyle cellStyle(GridGrindResponse.CellStyleReport style) {
    return new Assertion.CellStyle(style);
  }

  /** Returns one workbook-protection fact assertion. */
  public static Assertion.WorkbookProtectionFacts workbookProtection(
      WorkbookProtectionReport protection) {
    return new Assertion.WorkbookProtectionFacts(protection);
  }

  /** Returns one sheet-structure fact assertion. */
  public static Assertion.SheetStructureFacts sheetStructure(
      GridGrindResponse.SheetSummaryReport sheet) {
    return new Assertion.SheetStructureFacts(sheet);
  }

  /** Returns one named-range fact assertion. */
  public static Assertion.NamedRangeFacts namedRanges(
      List<GridGrindResponse.NamedRangeReport> namedRanges) {
    return new Assertion.NamedRangeFacts(namedRanges);
  }

  /** Returns one table fact assertion. */
  public static Assertion.TableFacts tables(List<TableEntryReport> tables) {
    return new Assertion.TableFacts(tables);
  }

  /** Returns one pivot-table fact assertion. */
  public static Assertion.PivotTableFacts pivotTables(List<PivotTableReport> pivotTables) {
    return new Assertion.PivotTableFacts(pivotTables);
  }

  /** Returns one chart fact assertion. */
  public static Assertion.ChartFacts charts(List<ChartReport> charts) {
    return new Assertion.ChartFacts(charts);
  }

  /** Returns one assertion that caps the maximum severity for one analysis query. */
  public static Assertion.AnalysisMaxSeverity analysisMaxSeverity(
      InspectionQuery.Analysis query, AnalysisSeverity maximumSeverity) {
    return new Assertion.AnalysisMaxSeverity(query, maximumSeverity);
  }

  /** Returns one assertion that requires a matching analysis finding. */
  public static Assertion.AnalysisFindingPresent analysisFindingPresent(
      InspectionQuery.Analysis query,
      AnalysisFindingCode code,
      AnalysisSeverity severity,
      String messageContains) {
    return new Assertion.AnalysisFindingPresent(query, code, severity, messageContains);
  }

  /** Returns one assertion that forbids a matching analysis finding. */
  public static Assertion.AnalysisFindingAbsent analysisFindingAbsent(
      InspectionQuery.Analysis query,
      AnalysisFindingCode code,
      AnalysisSeverity severity,
      String messageContains) {
    return new Assertion.AnalysisFindingAbsent(query, code, severity, messageContains);
  }

  /** Returns one composite assertion that requires every nested assertion to pass. */
  public static Assertion.AllOf allOf(Assertion... assertions) {
    return new Assertion.AllOf(List.of(assertions));
  }

  /** Returns one composite assertion that requires any nested assertion to pass. */
  public static Assertion.AnyOf anyOf(Assertion... assertions) {
    return new Assertion.AnyOf(List.of(assertions));
  }

  /** Returns one negated assertion. */
  public static Assertion.Not not(Assertion assertion) {
    return new Assertion.Not(assertion);
  }
}
