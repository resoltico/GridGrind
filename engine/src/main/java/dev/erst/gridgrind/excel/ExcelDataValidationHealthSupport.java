package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/** Derives workbook-health findings from canonical data-validation snapshots. */
final class ExcelDataValidationHealthSupport {
  List<WorkbookAnalysis.AnalysisFinding> dataValidationHealthFindings(
      String sheetName, List<ExcelDataValidationSnapshot> snapshots) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (ExcelDataValidationSnapshot snapshot : snapshots) {
      switch (snapshot) {
        case ExcelDataValidationSnapshot.Supported supported ->
            findings.addAll(supportedFindings(sheetName, supported));
        case ExcelDataValidationSnapshot.Unsupported unsupported ->
            findings.add(unsupportedFinding(sheetName, unsupported));
      }
    }
    findings.addAll(overlapFindings(sheetName, snapshots));
    return List.copyOf(findings);
  }

  private static List<WorkbookAnalysis.AnalysisFinding> supportedFindings(
      String sheetName, ExcelDataValidationSnapshot.Supported supported) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    WorkbookAnalysis.AnalysisLocation.Range location =
        new WorkbookAnalysis.AnalysisLocation.Range(sheetName, supported.ranges().getFirst());
    switch (supported.validation().rule()) {
      case ExcelDataValidationRule.ExplicitList explicitList -> {
        if (explicitList.values().isEmpty()) {
          findings.add(
              new WorkbookAnalysis.AnalysisFinding(
                  AnalysisFindingCode.DATA_VALIDATION_EMPTY_EXPLICIT_LIST,
                  AnalysisSeverity.ERROR,
                  "Explicit-list validation is empty",
                  "Explicit-list validation contains no values.",
                  location,
                  supported.ranges()));
        }
      }
      case ExcelDataValidationRule.FormulaList formulaList ->
          addBrokenFormulaFindingIfNeeded(
              findings, location, formulaList.formula(), "list validation formula");
      case ExcelDataValidationRule.WholeNumber wholeNumber -> {
        addBrokenFormulaFindingIfNeeded(
            findings, location, wholeNumber.formula1(), "whole-number validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, wholeNumber.formula2(), "whole-number validation formula");
      }
      case ExcelDataValidationRule.DecimalNumber decimalNumber -> {
        addBrokenFormulaFindingIfNeeded(
            findings, location, decimalNumber.formula1(), "decimal validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, decimalNumber.formula2(), "decimal validation formula");
      }
      case ExcelDataValidationRule.DateRule dateRule -> {
        addBrokenFormulaFindingIfNeeded(
            findings, location, dateRule.formula1(), "date validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, dateRule.formula2(), "date validation formula");
      }
      case ExcelDataValidationRule.TimeRule timeRule -> {
        addBrokenFormulaFindingIfNeeded(
            findings, location, timeRule.formula1(), "time validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, timeRule.formula2(), "time validation formula");
      }
      case ExcelDataValidationRule.TextLength textLength -> {
        addBrokenFormulaFindingIfNeeded(
            findings, location, textLength.formula1(), "text-length validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, textLength.formula2(), "text-length validation formula");
      }
      case ExcelDataValidationRule.CustomFormula customFormula ->
          addBrokenFormulaFindingIfNeeded(
              findings, location, customFormula.formula(), "custom validation formula");
    }
    return List.copyOf(findings);
  }

  private static WorkbookAnalysis.AnalysisFinding unsupportedFinding(
      String sheetName, ExcelDataValidationSnapshot.Unsupported unsupported) {
    WorkbookAnalysis.AnalysisLocation.Range location =
        new WorkbookAnalysis.AnalysisLocation.Range(sheetName, unsupported.ranges().getFirst());
    if ("MISSING_FORMULA".equals(unsupported.kind())) {
      return new WorkbookAnalysis.AnalysisFinding(
          AnalysisFindingCode.DATA_VALIDATION_MALFORMED_RULE,
          AnalysisSeverity.ERROR,
          "Data-validation rule is malformed",
          unsupported.detail(),
          location,
          unsupported.ranges());
    }
    return new WorkbookAnalysis.AnalysisFinding(
        AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
        AnalysisSeverity.WARNING,
        "Unsupported data-validation rule",
        unsupported.detail(),
        location,
        unsupported.ranges());
  }

  private static void addBrokenFormulaFindingIfNeeded(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      WorkbookAnalysis.AnalysisLocation.Range location,
      String formula,
      String label) {
    if (formula == null) {
      return;
    }
    if (formula.toUpperCase(Locale.ROOT).contains("#REF!")) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
              AnalysisSeverity.ERROR,
              "Broken data-validation formula",
              "Data-validation " + label + " contains a broken reference.",
              location,
              List.of(formula)));
    }
  }

  private static List<WorkbookAnalysis.AnalysisFinding> overlapFindings(
      String sheetName, List<ExcelDataValidationSnapshot> snapshots) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (int firstIndex = 0; firstIndex < snapshots.size(); firstIndex++) {
      for (int secondIndex = firstIndex + 1; secondIndex < snapshots.size(); secondIndex++) {
        for (String firstRange : snapshots.get(firstIndex).ranges()) {
          for (String secondRange : snapshots.get(secondIndex).ranges()) {
            if (!ExcelDataValidationRangeSupport.intersects(
                ExcelRange.parse(firstRange), ExcelRange.parse(secondRange))) {
              continue;
            }
            findings.add(
                new WorkbookAnalysis.AnalysisFinding(
                    AnalysisFindingCode.DATA_VALIDATION_OVERLAPPING_RULES,
                    AnalysisSeverity.WARNING,
                    "Overlapping data-validation rules",
                    "Two data-validation structures overlap on the same sheet.",
                    new WorkbookAnalysis.AnalysisLocation.Range(sheetName, firstRange),
                    List.of(firstRange, secondRange)));
          }
        }
      }
    }
    return List.copyOf(findings);
  }
}
