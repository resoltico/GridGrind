package dev.erst.gridgrind.excel.validation;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.excel.ExcelRange;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

/** Derives workbook-health findings from canonical data-validation snapshots. */
public final class ExcelDataValidationHealthSupport {
  List<AnalysisFindingReport> dataValidationHealthFindings(
      String sheetName, List<DataValidationEntryReport> snapshots) {
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (DataValidationEntryReport snapshot : snapshots) {
      switch (snapshot) {
        case DataValidationEntryReport.Supported supported ->
            findings.addAll(supportedFindings(sheetName, supported));
        case DataValidationEntryReport.Unsupported unsupported ->
            findings.add(unsupportedFinding(sheetName, unsupported));
      }
    }
    findings.addAll(overlapFindings(sheetName, snapshots));
    return List.copyOf(findings);
  }

  private static List<AnalysisFindingReport> supportedFindings(
      String sheetName, DataValidationEntryReport.Supported supported) {
    List<AnalysisFindingReport> findings = new ArrayList<>();
    AnalysisLocationReport.Range location =
        new AnalysisLocationReport.Range(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName),
            supported.ranges().getFirst().value());
    switch (supported.validation().rule()) {
      case DataValidationRuleSpec.ExplicitList explicitList -> {
        if (explicitList.values().isEmpty()) {
          findings.add(
              new AnalysisFindingReport(
                  AnalysisFindingCode.DATA_VALIDATION_EMPTY_EXPLICIT_LIST,
                  AnalysisSeverity.ERROR,
                  "Explicit-list validation is empty",
                  "Explicit-list validation contains no values.",
                  location,
                  supported.ranges().stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
                      .toList()));
        }
      }
      case DataValidationRuleSpec.FormulaList formulaList ->
          addBrokenFormulaFindingIfNeeded(
              findings, location, Optional.of(formulaList.formula()), "list validation formula");
      case DataValidationRuleSpec.WholeNumber wholeNumber -> {
        addBrokenFormulaFindingIfNeeded(
            findings,
            location,
            Optional.of(wholeNumber.formula1()),
            "whole-number validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, wholeNumber.formula2(), "whole-number validation formula");
      }
      case DataValidationRuleSpec.DecimalNumber decimalNumber -> {
        addBrokenFormulaFindingIfNeeded(
            findings,
            location,
            Optional.of(decimalNumber.formula1()),
            "decimal validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, decimalNumber.formula2(), "decimal validation formula");
      }
      case DataValidationRuleSpec.DateRule dateRule -> {
        addBrokenFormulaFindingIfNeeded(
            findings, location, Optional.of(dateRule.formula1()), "date validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, dateRule.formula2(), "date validation formula");
      }
      case DataValidationRuleSpec.TimeRule timeRule -> {
        addBrokenFormulaFindingIfNeeded(
            findings, location, Optional.of(timeRule.formula1()), "time validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, timeRule.formula2(), "time validation formula");
      }
      case DataValidationRuleSpec.TextLength textLength -> {
        addBrokenFormulaFindingIfNeeded(
            findings,
            location,
            Optional.of(textLength.formula1()),
            "text-length validation formula");
        addBrokenFormulaFindingIfNeeded(
            findings, location, textLength.formula2(), "text-length validation formula");
      }
      case DataValidationRuleSpec.CustomFormula customFormula ->
          addBrokenFormulaFindingIfNeeded(
              findings,
              location,
              Optional.of(customFormula.formula()),
              "custom validation formula");
    }
    return List.copyOf(findings);
  }

  private static AnalysisFindingReport unsupportedFinding(
      String sheetName, DataValidationEntryReport.Unsupported unsupported) {
    AnalysisLocationReport.Range location =
        new AnalysisLocationReport.Range(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName),
            unsupported.ranges().getFirst().value());
    if ("MISSING_FORMULA".equals(unsupported.kind())) {
      return new AnalysisFindingReport(
          AnalysisFindingCode.DATA_VALIDATION_MALFORMED_RULE,
          AnalysisSeverity.ERROR,
          "Data-validation rule is malformed",
          unsupported.detail(),
          location,
          unsupported.ranges().stream()
              .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
              .toList());
    }
    return new AnalysisFindingReport(
        AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
        AnalysisSeverity.WARNING,
        "Unsupported data-validation rule",
        unsupported.detail(),
        location,
        unsupported.ranges().stream()
            .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
            .toList());
  }

  private static void addBrokenFormulaFindingIfNeeded(
      List<AnalysisFindingReport> findings,
      AnalysisLocationReport.Range location,
      Optional<String> formula,
      String label) {
    if (formula.isEmpty()) {
      return;
    }
    String actualFormula = formula.orElseThrow();
    if (actualFormula.toUpperCase(Locale.ROOT).contains("#REF!")) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
              AnalysisSeverity.ERROR,
              "Broken data-validation formula",
              "Data-validation " + label + " contains a broken reference.",
              location,
              List.of(actualFormula)));
    }
  }

  private static List<AnalysisFindingReport> overlapFindings(
      String sheetName, List<DataValidationEntryReport> snapshots) {
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (int firstIndex = 0; firstIndex < snapshots.size(); firstIndex++) {
      for (int secondIndex = firstIndex + 1; secondIndex < snapshots.size(); secondIndex++) {
        for (String firstRange :
            snapshots.get(firstIndex).ranges().stream()
                .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
                .toList()) {
          for (String secondRange :
              snapshots.get(secondIndex).ranges().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
                  .toList()) {
            if (!ExcelDataValidationRangeSupport.intersects(
                ExcelRange.parse(firstRange), ExcelRange.parse(secondRange))) {
              continue;
            }
            findings.add(
                new AnalysisFindingReport(
                    AnalysisFindingCode.DATA_VALIDATION_OVERLAPPING_RULES,
                    AnalysisSeverity.WARNING,
                    "Overlapping data-validation rules",
                    "Two data-validation structures overlap on the same sheet.",
                    new AnalysisLocationReport.Range(
                        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName), firstRange),
                    List.of(firstRange, secondRange)));
          }
        }
      }
    }
    return List.copyOf(findings);
  }
}
