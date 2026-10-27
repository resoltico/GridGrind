package dev.erst.gridgrind.engine.validation;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.engine.formulas.ComparisonFormulas;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import dev.erst.gridgrind.engine.poi.SqrefRanges;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import javax.xml.namespace.QName;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaShifter;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Owns data-validation writes, truthful readback, range rewrites and diagnostics. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
public final class DataValidationAdapter {
  /** Returns the count of stored validation elements, including unsupported ones. */
  public int dataValidationCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    return dataValidations == null ? 0 : dataValidations.sizeOfDataValidationArray();
  }

  /** Derives diagnostics from factual validation readback. */
  public List<AnalysisFindingReport> dataValidationHealthFindings(
      String sheetName, XSSFSheet sheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    return analyzeValidations(sheetName, dataValidations(sheet));
  }

  public void setDataValidation(
      XSSFSheet sheet, String range, ResolvedDataValidation validationDefinition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(validationDefinition, "validationDefinition must not be null");
    CellBounds targetRange = CellBounds.parse(range);
    DataValidationHelper helper = sheet.getDataValidationHelper();
    DataValidationConstraint constraint = toConstraint(helper, validationDefinition.rule());
    normalizeOverlappingSqref(sheet, List.of(targetRange));
    CellRangeAddressList regions =
        new CellRangeAddressList(
            targetRange.firstRow(),
            targetRange.lastRow(),
            targetRange.firstColumn(),
            targetRange.lastColumn());
    DataValidation validation = helper.createValidation(constraint, regions);
    validation.setEmptyCellAllowed(validationDefinition.allowBlank());
    validation.setSuppressDropDownArrow(validationDefinition.suppressDropDownArrow());
    if (validationDefinition.prompt().isEmpty()) {
      validation.setShowPromptBox(false);
    } else {
      ResolvedValidationPrompt prompt = validationDefinition.prompt().orElseThrow();
      validation.setShowPromptBox(prompt.showPromptBox());
      validation.createPromptBox(prompt.title(), prompt.text());
    }
    if (validationDefinition.errorAlert().isEmpty()) {
      validation.setShowErrorBox(false);
    } else {
      ResolvedValidationErrorAlert errorAlert = validationDefinition.errorAlert().orElseThrow();
      validation.setShowErrorBox(errorAlert.showErrorBox());
      validation.setErrorStyle(toPoiErrorStyle(errorAlert.style()));
      validation.createErrorBox(errorAlert.title(), errorAlert.text());
    }
    sheet.addValidationData(validation);
    syncValidationCount(sheet);
  }

  public void clearDataValidations(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (sheet.getCTWorksheet().isSetDataValidations())
      sheet.getCTWorksheet().unsetDataValidations();
  }

  public void clearDataValidations(
      XSSFSheet sheet, List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    normalizeOverlappingSqref(
        sheet, ranges.stream().map(range -> CellBounds.parse(range.value())).toList());
  }

  public static DataValidationConstraint toConstraint(
      DataValidationHelper helper, DataValidationRuleSpec rule) {
    return switch (rule) {
      case DataValidationRuleSpec.ExplicitList explicitList ->
          explicitList.values().isEmpty()
              ? helper.createFormulaListConstraint("\"\"")
              : helper.createExplicitListConstraint(explicitList.values().toArray(String[]::new));
      case DataValidationRuleSpec.FormulaList formulaList ->
          helper.createFormulaListConstraint(
              ComparisonFormulas.normalizeFormula(formulaList.formula(), "formula"));
      case DataValidationRuleSpec.WholeNumber wholeNumber ->
          helper.createIntegerConstraint(
              toPoi(wholeNumber.operator()),
              ComparisonFormulas.normalizeFormula(wholeNumber.formula1(), "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(wholeNumber.formula2(), "formula2")
                  .orElse(null));
      case DataValidationRuleSpec.DecimalNumber decimalNumber ->
          helper.createDecimalConstraint(
              toPoi(decimalNumber.operator()),
              ComparisonFormulas.normalizeFormula(decimalNumber.formula1(), "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(decimalNumber.formula2(), "formula2")
                  .orElse(null));
      case DataValidationRuleSpec.DateRule dateRule ->
          helper.createDateConstraint(
              toPoi(dateRule.operator()),
              ComparisonFormulas.normalizeFormula(dateRule.formula1(), "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(dateRule.formula2(), "formula2")
                  .orElse(null),
              null);
      case DataValidationRuleSpec.TimeRule timeRule ->
          helper.createTimeConstraint(
              toPoi(timeRule.operator()),
              ComparisonFormulas.normalizeFormula(timeRule.formula1(), "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(timeRule.formula2(), "formula2")
                  .orElse(null));
      case DataValidationRuleSpec.TextLength textLength ->
          helper.createTextLengthConstraint(
              toPoi(textLength.operator()),
              ComparisonFormulas.normalizeFormula(textLength.formula1(), "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(textLength.formula2(), "formula2")
                  .orElse(null));
      case DataValidationRuleSpec.CustomFormula customFormula ->
          helper.createCustomConstraint(
              ComparisonFormulas.normalizeFormula(customFormula.formula(), "formula"));
    };
  }

  public List<DataValidationEntryReport> dataValidations(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return readDataValidations(sheet, ranges -> true);
  }

  public List<DataValidationEntryReport> dataValidations(
      XSSFSheet sheet, List<dev.erst.gridgrind.core.vocabulary.RangeRef> targets) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    var copied = List.copyOf(targets);
    return readDataValidations(sheet, ranges -> matchesSelection(ranges, copied));
  }

  private List<DataValidationEntryReport> readDataValidations(
      XSSFSheet sheet, java.util.function.Predicate<List<String>> matches) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    if (dataValidations == null) {
      return List.of();
    }
    List<DataValidationEntryReport> snapshots = new ArrayList<>();
    for (CTDataValidation validation : dataValidations.getDataValidationArray()) {
      List<String> ranges = SqrefRanges.normalizedSqref(validation.getSqref());
      if (!matches.test(ranges)) {
        continue;
      }
      snapshots.add(toSnapshot(validation, ranges));
    }
    return List.copyOf(snapshots);
  }

  public static DataValidationEntryReport toSnapshot(
      XSSFDataValidation validation, List<String> ranges) {
    DataValidationConstraint constraint = validation.getValidationConstraint();
    return switch (constraint.getValidationType()) {
      case DataValidationConstraint.ValidationType.ANY ->
          new DataValidationEntryReport.Unsupported(
              ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
              "ANY",
              "Excel 'any value' validation is not modeled by GridGrind.");
      case DataValidationConstraint.ValidationType.LIST -> listSnapshot(validation, ranges);
      case DataValidationConstraint.ValidationType.INTEGER ->
          comparisonSnapshot(validation, ranges, constraint, ComparisonFamily.INTEGER);
      case DataValidationConstraint.ValidationType.DECIMAL ->
          comparisonSnapshot(validation, ranges, constraint, ComparisonFamily.DECIMAL);
      case DataValidationConstraint.ValidationType.DATE ->
          comparisonSnapshot(validation, ranges, constraint, ComparisonFamily.DATE);
      case DataValidationConstraint.ValidationType.TIME ->
          comparisonSnapshot(validation, ranges, constraint, ComparisonFamily.TIME);
      case DataValidationConstraint.ValidationType.TEXT_LENGTH ->
          comparisonSnapshot(validation, ranges, constraint, ComparisonFamily.TEXT_LENGTH);
      case DataValidationConstraint.ValidationType.FORMULA ->
          customFormulaSnapshot(validation, ranges);
      default ->
          new DataValidationEntryReport.Unsupported(
              ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
              "UNKNOWN",
              "Unsupported data-validation type: " + constraint.getValidationType());
    };
  }

  public static DataValidationEntryReport toSnapshot(
      CTDataValidation validation, List<String> ranges) {
    String type = Objects.requireNonNullElse(rawAttribute(validation, "type"), "");
    return switch (type.toLowerCase(Locale.ROOT)) {
      case "none", "" ->
          new DataValidationEntryReport.Unsupported(
              ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
              "ANY",
              "Excel 'any value' validation is not modeled by GridGrind.");
      case "list" -> listSnapshot(validation, ranges);
      case "whole" -> comparisonSnapshot(validation, ranges, ComparisonFamily.INTEGER);
      case "decimal" -> comparisonSnapshot(validation, ranges, ComparisonFamily.DECIMAL);
      case "date" -> comparisonSnapshot(validation, ranges, ComparisonFamily.DATE);
      case "time" -> comparisonSnapshot(validation, ranges, ComparisonFamily.TIME);
      case "textlength" -> comparisonSnapshot(validation, ranges, ComparisonFamily.TEXT_LENGTH);
      case "custom" -> customFormulaSnapshot(validation, ranges);
      default ->
          new DataValidationEntryReport.Unsupported(
              ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
              "UNKNOWN",
              "Unsupported data-validation type: " + Objects.requireNonNullElse(type, "UNKNOWN"));
    };
  }

  private static DataValidationEntryReport listSnapshot(
      XSSFDataValidation validation, List<String> ranges) {
    DataValidationConstraint constraint = validation.getValidationConstraint();
    String[] explicitValues = constraint.getExplicitListValues();
    if (explicitValues != null) {
      List<String> explicitListValues =
          "\"\"".equals(Objects.requireNonNullElse(constraint.getFormula1(), ""))
              ? List.of()
              : List.of(explicitValues);
      return new DataValidationEntryReport.Supported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          new DataValidationEntryReport.DataValidationDefinitionReport(
              new DataValidationRuleSpec.ExplicitList(explicitListValues),
              validation.getEmptyCellAllowed(),
              validation.getSuppressDropDownArrow(),
              prompt(validation),
              errorAlert(validation)));
    }
    String formula1 = constraint.getFormula1();
    if (formula1 == null || formula1.isBlank()) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "MISSING_FORMULA",
          "List validation is missing both explicit values and formula1.");
    }
    return new DataValidationEntryReport.Supported(
        ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
        new DataValidationEntryReport.DataValidationDefinitionReport(
            new DataValidationRuleSpec.FormulaList(
                ComparisonFormulas.normalizeFormula(formula1, "formula")),
            validation.getEmptyCellAllowed(),
            validation.getSuppressDropDownArrow(),
            prompt(validation),
            errorAlert(validation)));
  }

  private static DataValidationEntryReport listSnapshot(
      CTDataValidation validation, List<String> ranges) {
    String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
    if (formula1.isBlank()) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "MISSING_FORMULA",
          "List validation is missing both explicit values and formula1.");
    }
    DataValidationRuleSpec rule;
    if ("\"\"".equals(formula1)) {
      rule = new DataValidationRuleSpec.ExplicitList(List.of());
    } else if (formula1.length() >= 2 && formula1.startsWith("\"") && formula1.endsWith("\"")) {
      String rawValues = formula1.substring(1, formula1.length() - 1);
      rule = new DataValidationRuleSpec.ExplicitList(Arrays.asList(rawValues.split(",")));
    } else {
      rule =
          new DataValidationRuleSpec.FormulaList(
              ComparisonFormulas.normalizeFormula(formula1, "formula"));
    }
    return new DataValidationEntryReport.Supported(
        ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
        new DataValidationEntryReport.DataValidationDefinitionReport(
            rule,
            validation.isSetAllowBlank() && validation.getAllowBlank(),
            suppressDropDownArrow(validation),
            prompt(validation),
            errorAlert(validation)));
  }

  private static DataValidationEntryReport comparisonSnapshot(
      XSSFDataValidation validation,
      List<String> ranges,
      DataValidationConstraint constraint,
      ComparisonFamily family) {
    String formula1 = constraint.getFormula1();
    if (formula1 == null || formula1.isBlank()) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "MISSING_FORMULA",
          family.label + " validation is missing formula1.");
    }
    ExcelComparisonOperator operator;
    try {
      operator = fromPoi(constraint.getOperator());
    } catch (IllegalArgumentException exception) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "UNKNOWN_OPERATOR",
          family.label + " validation uses an unsupported comparison operator.");
    }
    Optional<String> formula2 = Optional.ofNullable(constraint.getFormula2());
    if (requiresUpperBound(operator) && (formula2.isEmpty() || formula2.orElseThrow().isBlank())) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "MISSING_FORMULA",
          family.label + " validation is missing formula2 for " + operatorLabel(operator) + ".");
    }
    DataValidationRuleSpec rule = comparisonRule(family, operator, formula1, formula2);
    return new DataValidationEntryReport.Supported(
        ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
        new DataValidationEntryReport.DataValidationDefinitionReport(
            rule,
            validation.getEmptyCellAllowed(),
            validation.getSuppressDropDownArrow(),
            prompt(validation),
            errorAlert(validation)));
  }

  private static DataValidationEntryReport comparisonSnapshot(
      CTDataValidation validation, List<String> ranges, ComparisonFamily family) {
    String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
    if (formula1.isBlank()) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "MISSING_FORMULA",
          family.label + " validation is missing formula1.");
    }
    ExcelComparisonOperator operator;
    try {
      operator = comparisonOperator(validation);
    } catch (IllegalArgumentException exception) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "UNKNOWN_OPERATOR",
          family.label + " validation uses an unsupported comparison operator.");
    }
    Optional<String> formula2 =
        validation.isSetFormula2() ? Optional.of(validation.getFormula2()) : Optional.empty();
    if (requiresUpperBound(operator) && (formula2.isEmpty() || formula2.orElseThrow().isBlank())) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "MISSING_FORMULA",
          family.label + " validation is missing formula2 for " + operatorLabel(operator) + ".");
    }
    DataValidationRuleSpec rule = comparisonRule(family, operator, formula1, formula2);
    return new DataValidationEntryReport.Supported(
        ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
        new DataValidationEntryReport.DataValidationDefinitionReport(
            rule,
            validation.isSetAllowBlank() && validation.getAllowBlank(),
            suppressDropDownArrow(validation),
            prompt(validation),
            errorAlert(validation)));
  }

  private static DataValidationEntryReport customFormulaSnapshot(
      XSSFDataValidation validation, List<String> ranges) {
    String formula1 = validation.getValidationConstraint().getFormula1();
    if (formula1 == null || formula1.isBlank()) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "MISSING_FORMULA",
          "Custom formula validation is missing formula1.");
    }
    return new DataValidationEntryReport.Supported(
        ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
        new DataValidationEntryReport.DataValidationDefinitionReport(
            new DataValidationRuleSpec.CustomFormula(
                ComparisonFormulas.normalizeFormula(formula1, "formula")),
            validation.getEmptyCellAllowed(),
            validation.getSuppressDropDownArrow(),
            prompt(validation),
            errorAlert(validation)));
  }

  private static DataValidationEntryReport customFormulaSnapshot(
      CTDataValidation validation, List<String> ranges) {
    String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
    if (formula1.isBlank()) {
      return new DataValidationEntryReport.Unsupported(
          ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
          "MISSING_FORMULA",
          "Custom formula validation is missing formula1.");
    }
    return new DataValidationEntryReport.Supported(
        ranges.stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
        new DataValidationEntryReport.DataValidationDefinitionReport(
            new DataValidationRuleSpec.CustomFormula(
                ComparisonFormulas.normalizeFormula(formula1, "formula")),
            validation.isSetAllowBlank() && validation.getAllowBlank(),
            suppressDropDownArrow(validation),
            prompt(validation),
            errorAlert(validation)));
  }

  public static java.util.Optional<DataValidationPromptReport> prompt(
      XSSFDataValidation validation) {
    String title = validation.getPromptBoxTitle();
    String text = validation.getPromptBoxText();
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return java.util.Optional.empty();
    }
    return java.util.Optional.of(
        new DataValidationPromptReport(title, text, validation.getShowPromptBox()));
  }

  public static java.util.Optional<DataValidationPromptReport> prompt(CTDataValidation validation) {
    String title = validation.isSetPromptTitle() ? validation.getPromptTitle() : null;
    String text = validation.isSetPrompt() ? validation.getPrompt() : null;
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return java.util.Optional.empty();
    }
    return java.util.Optional.of(
        new DataValidationPromptReport(
            title, text, validation.isSetShowInputMessage() && validation.getShowInputMessage()));
  }

  public static java.util.Optional<DataValidationErrorAlertReport> errorAlert(
      XSSFDataValidation validation) {
    String title = validation.getErrorBoxTitle();
    String text = validation.getErrorBoxText();
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return java.util.Optional.empty();
    }
    return java.util.Optional.of(
        new DataValidationErrorAlertReport(
            fromPoiErrorStyle(validation.getErrorStyle()),
            title,
            text,
            validation.getShowErrorBox()));
  }

  public static java.util.Optional<DataValidationErrorAlertReport> errorAlert(
      CTDataValidation validation) {
    String title = validation.isSetErrorTitle() ? validation.getErrorTitle() : null;
    String text = validation.isSetError() ? validation.getError() : null;
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return java.util.Optional.empty();
    }
    return java.util.Optional.of(
        new DataValidationErrorAlertReport(
            errorStyle(validation),
            title,
            text,
            validation.isSetShowErrorMessage() && validation.getShowErrorMessage()));
  }

  public static ExcelComparisonOperator comparisonOperator(CTDataValidation validation) {
    String operator = Objects.requireNonNullElse(rawAttribute(validation, "operator"), "between");
    return fromXml(operator);
  }

  public static ExcelDataValidationErrorStyle errorStyle(CTDataValidation validation) {
    String errorStyle = Objects.requireNonNullElse(rawAttribute(validation, "errorStyle"), "stop");
    return switch (errorStyle) {
      case "stop" -> ExcelDataValidationErrorStyle.STOP;
      case "warning" -> ExcelDataValidationErrorStyle.WARNING;
      case "information" -> ExcelDataValidationErrorStyle.INFORMATION;
      default ->
          throw new IllegalArgumentException(
              "Unsupported raw validation error style: " + errorStyle);
    };
  }

  public static String operatorLabel(ExcelComparisonOperator operator) {
    return switch (operator) {
      case BETWEEN -> "between operator";
      case NOT_BETWEEN -> "not-between operator";
      case EQUAL -> "equal operator";
      case NOT_EQUAL -> "not-equal operator";
      case GREATER_THAN -> "greater-than operator";
      case GREATER_OR_EQUAL -> "greater-or-equal operator";
      case LESS_THAN -> "less-than operator";
      case LESS_OR_EQUAL -> "less-or-equal operator";
    };
  }

  private static DataValidationRuleSpec comparisonRule(
      ComparisonFamily family,
      ExcelComparisonOperator operator,
      String formula1,
      Optional<String> formula2) {
    return switch (family) {
      case INTEGER ->
          new DataValidationRuleSpec.WholeNumber(
              operator,
              ComparisonFormulas.normalizeFormula(formula1, "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(formula2, "formula2"));
      case DECIMAL ->
          new DataValidationRuleSpec.DecimalNumber(
              operator,
              ComparisonFormulas.normalizeFormula(formula1, "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(formula2, "formula2"));
      case DATE ->
          new DataValidationRuleSpec.DateRule(
              operator,
              ComparisonFormulas.normalizeFormula(formula1, "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(formula2, "formula2"));
      case TIME ->
          new DataValidationRuleSpec.TimeRule(
              operator,
              ComparisonFormulas.normalizeFormula(formula1, "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(formula2, "formula2"));
      case TEXT_LENGTH ->
          new DataValidationRuleSpec.TextLength(
              operator,
              ComparisonFormulas.normalizeFormula(formula1, "formula1"),
              ComparisonFormulas.normalizeOptionalFormula(formula2, "formula2"));
    };
  }

  private static boolean suppressDropDownArrow(CTDataValidation validation) {
    return !validation.isSetShowDropDown() || !validation.getShowDropDown();
  }

  private static String rawAttribute(CTDataValidation validation, String attributeName) {
    try (var cursor = validation.newCursor()) {
      return cursor.getAttributeText(new QName("", attributeName));
    }
  }

  private static boolean requiresUpperBound(ExcelComparisonOperator operator) {
    return operator == ExcelComparisonOperator.BETWEEN
        || operator == ExcelComparisonOperator.NOT_BETWEEN;
  }

  /** Groups comparison-rule families so validation diagnostics stay precise. */
  enum ComparisonFamily {
    INTEGER("whole-number"),
    DECIMAL("decimal"),
    DATE("date"),
    TIME("time"),
    TEXT_LENGTH("text-length");

    private final String label;

    ComparisonFamily(String label) {
      this.label = label;
    }
  }

  private List<AnalysisFindingReport> analyzeValidations(
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
            if (!SheetStructure.intersects(
                CellBounds.parse(firstRange), CellBounds.parse(secondRange))) {
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

  public static void normalizeOverlappingSqref(XSSFSheet sheet, List<CellBounds> cutouts) {
    if (!sheet.getCTWorksheet().isSetDataValidations()) {
      return;
    }
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    for (int validationIndex = dataValidations.sizeOfDataValidationArray() - 1;
        validationIndex >= 0;
        validationIndex--) {
      CTDataValidation validation = dataValidations.getDataValidationArray(validationIndex);
      List<String> retainedRanges = retainedRanges(validation, cutouts);
      if (retainedRanges.isEmpty()) {
        dataValidations.removeDataValidation(validationIndex);
        continue;
      }
      validation.setSqref(retainedRanges);
    }
    syncValidationCount(sheet);
  }

  public static List<String> retainedRanges(CTDataValidation validation, List<CellBounds> cutouts) {
    List<CellBounds> retained =
        SqrefRanges.normalizedSqref(validation.getSqref()).stream().map(CellBounds::parse).toList();
    List<CellBounds> next = new ArrayList<>();
    for (CellBounds cutout : cutouts) {
      next.clear();
      for (CellBounds existingRange : retained) {
        next.addAll(subtract(existingRange, cutout));
      }
      retained = List.copyOf(next);
    }
    return retained.stream().map(SheetStructure::formatRange).toList();
  }

  public static List<CellBounds> subtract(CellBounds source, CellBounds cutout) {
    if (!SheetStructure.intersects(source, cutout)) {
      return List.of(source);
    }
    int intersectionFirstRow = Math.max(source.firstRow(), cutout.firstRow());
    int intersectionLastRow = Math.min(source.lastRow(), cutout.lastRow());
    int intersectionFirstColumn = Math.max(source.firstColumn(), cutout.firstColumn());
    int intersectionLastColumn = Math.min(source.lastColumn(), cutout.lastColumn());

    List<CellBounds> retained = new ArrayList<>(4);
    if (source.firstRow() < intersectionFirstRow) {
      retained.add(
          new CellBounds(
              source.firstRow(),
              intersectionFirstRow - 1,
              source.firstColumn(),
              source.lastColumn()));
    }
    if (intersectionLastRow < source.lastRow()) {
      retained.add(
          new CellBounds(
              intersectionLastRow + 1,
              source.lastRow(),
              source.firstColumn(),
              source.lastColumn()));
    }
    if (source.firstColumn() < intersectionFirstColumn) {
      retained.add(
          new CellBounds(
              intersectionFirstRow,
              intersectionLastRow,
              source.firstColumn(),
              intersectionFirstColumn - 1));
    }
    if (intersectionLastColumn < source.lastColumn()) {
      retained.add(
          new CellBounds(
              intersectionFirstRow,
              intersectionLastRow,
              intersectionLastColumn + 1,
              source.lastColumn()));
    }
    return List.copyOf(retained);
  }

  public static void syncValidationCount(XSSFSheet sheet) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    int count = dataValidations.sizeOfDataValidationArray();
    if (count == 0) {
      sheet.getCTWorksheet().unsetDataValidations();
      return;
    }
    dataValidations.setCount(count);
  }

  public static boolean matchesSelection(
      List<String> ranges, List<dev.erst.gridgrind.core.vocabulary.RangeRef> targets) {
    return ranges.stream()
        .map(CellBounds::parse)
        .anyMatch(
            existing ->
                targets.stream()
                    .map(target -> CellBounds.parse(target.value()))
                    .anyMatch(target -> SheetStructure.intersects(existing, target)));
  }

  public static int toPoiErrorStyle(ExcelDataValidationErrorStyle errorStyle) {
    return switch (errorStyle) {
      case STOP -> DataValidation.ErrorStyle.STOP;
      case WARNING -> DataValidation.ErrorStyle.WARNING;
      case INFORMATION -> DataValidation.ErrorStyle.INFO;
    };
  }

  public static ExcelDataValidationErrorStyle fromPoiErrorStyle(int errorStyle) {
    return switch (errorStyle) {
      case DataValidation.ErrorStyle.STOP -> ExcelDataValidationErrorStyle.STOP;
      case DataValidation.ErrorStyle.WARNING -> ExcelDataValidationErrorStyle.WARNING;
      case DataValidation.ErrorStyle.INFO -> ExcelDataValidationErrorStyle.INFORMATION;
      default ->
          throw new IllegalArgumentException("Unsupported Apache POI error style: " + errorStyle);
    };
  }

  public static int toPoi(ExcelComparisonOperator operator) {
    return switch (operator) {
      case BETWEEN -> DataValidationConstraint.OperatorType.BETWEEN;
      case NOT_BETWEEN -> DataValidationConstraint.OperatorType.NOT_BETWEEN;
      case EQUAL -> DataValidationConstraint.OperatorType.EQUAL;
      case NOT_EQUAL -> DataValidationConstraint.OperatorType.NOT_EQUAL;
      case GREATER_THAN -> DataValidationConstraint.OperatorType.GREATER_THAN;
      case LESS_THAN -> DataValidationConstraint.OperatorType.LESS_THAN;
      case GREATER_OR_EQUAL -> DataValidationConstraint.OperatorType.GREATER_OR_EQUAL;
      case LESS_OR_EQUAL -> DataValidationConstraint.OperatorType.LESS_OR_EQUAL;
    };
  }

  public static ExcelComparisonOperator fromPoi(int operator) {
    return switch (operator) {
      case DataValidationConstraint.OperatorType.BETWEEN -> ExcelComparisonOperator.BETWEEN;
      case DataValidationConstraint.OperatorType.NOT_BETWEEN -> ExcelComparisonOperator.NOT_BETWEEN;
      case DataValidationConstraint.OperatorType.EQUAL -> ExcelComparisonOperator.EQUAL;
      case DataValidationConstraint.OperatorType.NOT_EQUAL -> ExcelComparisonOperator.NOT_EQUAL;
      case DataValidationConstraint.OperatorType.GREATER_THAN ->
          ExcelComparisonOperator.GREATER_THAN;
      case DataValidationConstraint.OperatorType.LESS_THAN -> ExcelComparisonOperator.LESS_THAN;
      case DataValidationConstraint.OperatorType.GREATER_OR_EQUAL ->
          ExcelComparisonOperator.GREATER_OR_EQUAL;
      case DataValidationConstraint.OperatorType.LESS_OR_EQUAL ->
          ExcelComparisonOperator.LESS_OR_EQUAL;
      default ->
          throw new IllegalArgumentException(
              "Unsupported Apache POI data-validation comparison operator: " + operator);
    };
  }

  public static ExcelComparisonOperator fromXml(String operator) {
    return switch (operator) {
      case "between" -> ExcelComparisonOperator.BETWEEN;
      case "notBetween" -> ExcelComparisonOperator.NOT_BETWEEN;
      case "equal" -> ExcelComparisonOperator.EQUAL;
      case "notEqual" -> ExcelComparisonOperator.NOT_EQUAL;
      case "greaterThan" -> ExcelComparisonOperator.GREATER_THAN;
      case "lessThan" -> ExcelComparisonOperator.LESS_THAN;
      case "greaterThanOrEqual" -> ExcelComparisonOperator.GREATER_OR_EQUAL;
      case "lessThanOrEqual" -> ExcelComparisonOperator.LESS_OR_EQUAL;
      default ->
          throw new IllegalArgumentException(
              "Unsupported OOXML data-validation comparison operator: " + operator);
    };
  }

  @SuppressWarnings("PMD.CloseResource")
  public static List<CTDataValidation> expectedValidationsAfterInsertRows(
      XSSFSheet sheet, int rowIndex, int rowCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getCTWorksheet().isSetDataValidations()) {
      return List.of();
    }
    XSSFWorkbook workbook = sheet.getWorkbook();
    int sheetIndex = workbook.getSheetIndex(sheet);
    XSSFEvaluationWorkbook evaluationWorkbook = XSSFEvaluationWorkbook.create(workbook);
    FormulaShifter formulaShifter =
        FormulaShifter.createForRowShift(
            sheetIndex,
            sheet.getSheetName(),
            rowIndex,
            SpreadsheetVersion.EXCEL2007.getLastRowIndex(),
            rowCount,
            SpreadsheetVersion.EXCEL2007);
    List<CTDataValidation> rewritten = new ArrayList<>();
    for (CTDataValidation validation : copiedValidations(sheet)) {
      List<String> shiftedRanges =
          shiftedRangesForInsertedRows(sheet, validation, rowIndex, rowCount);
      if (shiftedRanges.isEmpty()) {
        continue;
      }
      validation.setSqref(shiftedRanges);
      shiftValidationFormulas(validation, evaluationWorkbook, formulaShifter, sheetIndex, sheet);
      rewritten.add(validation);
    }
    return List.copyOf(rewritten);
  }

  @SuppressWarnings("PMD.CloseResource")
  public static List<CTDataValidation> expectedValidationsAfterInsertColumns(
      XSSFSheet sheet, int columnIndex, int columnCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getCTWorksheet().isSetDataValidations()) {
      return List.of();
    }
    XSSFWorkbook workbook = sheet.getWorkbook();
    int sheetIndex = workbook.getSheetIndex(sheet);
    XSSFEvaluationWorkbook evaluationWorkbook = XSSFEvaluationWorkbook.create(workbook);
    FormulaShifter formulaShifter =
        FormulaShifter.createForColumnShift(
            sheetIndex,
            sheet.getSheetName(),
            columnIndex,
            SpreadsheetVersion.EXCEL2007.getLastColumnIndex(),
            columnCount,
            SpreadsheetVersion.EXCEL2007);
    List<CTDataValidation> rewritten = new ArrayList<>();
    for (CTDataValidation validation : copiedValidations(sheet)) {
      List<String> shiftedRanges =
          shiftedRangesForInsertedColumns(sheet, validation, columnIndex, columnCount);
      if (shiftedRanges.isEmpty()) {
        continue;
      }
      validation.setSqref(shiftedRanges);
      shiftValidationFormulas(validation, evaluationWorkbook, formulaShifter, sheetIndex, sheet);
      rewritten.add(validation);
    }
    return List.copyOf(rewritten);
  }

  public static void replaceDataValidations(XSSFSheet sheet, List<CTDataValidation> validations) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(validations, "validations must not be null");
    if (sheet.getCTWorksheet().isSetDataValidations()) {
      sheet.getCTWorksheet().unsetDataValidations();
    }
    if (validations.isEmpty()) {
      return;
    }
    CTDataValidations target = sheet.getCTWorksheet().addNewDataValidations();
    for (CTDataValidation validation : validations) {
      CTDataValidation copied = target.addNewDataValidation();
      copied.set(validation);
    }
    target.setCount(target.sizeOfDataValidationArray());
  }

  private static List<CTDataValidation> copiedValidations(XSSFSheet sheet) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    List<CTDataValidation> copied = new ArrayList<>(dataValidations.sizeOfDataValidationArray());
    for (CTDataValidation validation : dataValidations.getDataValidationArray()) {
      copied.add((CTDataValidation) validation.copy());
    }
    return List.copyOf(copied);
  }

  private static List<String> shiftedRangesForInsertedRows(
      XSSFSheet sheet, CTDataValidation validation, int rowIndex, int rowCount) {
    List<String> shifted = new ArrayList<>();
    for (String rawRange : SqrefRanges.normalizedSqref(validation.getSqref())) {
      CellBounds range = CellBounds.parse(rawRange);
      shifted.addAll(shiftedRangesForInsertedRows(sheet, rawRange, range, rowIndex, rowCount));
    }
    return List.copyOf(shifted);
  }

  private static List<String> shiftedRangesForInsertedColumns(
      XSSFSheet sheet, CTDataValidation validation, int columnIndex, int columnCount) {
    List<String> shifted = new ArrayList<>();
    for (String rawRange : SqrefRanges.normalizedSqref(validation.getSqref())) {
      CellBounds range = CellBounds.parse(rawRange);
      shifted.addAll(
          shiftedRangesForInsertedColumns(sheet, rawRange, range, columnIndex, columnCount));
    }
    return List.copyOf(shifted);
  }

  private static List<String> shiftedRangesForInsertedRows(
      XSSFSheet sheet, String rawRange, CellBounds range, int rowIndex, int rowCount) {
    if (range.lastRow() < rowIndex) {
      return List.of(rawRange);
    }
    if (range.firstRow() >= rowIndex) {
      return List.of(
          shiftedRangeForRows(
              sheet,
              rawRange,
              range.firstRow() + rowCount,
              range.lastRow() + rowCount,
              range.firstColumn(),
              range.lastColumn()));
    }
    List<String> shifted = new ArrayList<>(2);
    shifted.add(
        SheetStructure.formatRange(
            new CellBounds(
                range.firstRow(), rowIndex - 1, range.firstColumn(), range.lastColumn())));
    shifted.add(
        shiftedRangeForRows(
            sheet,
            rawRange,
            rowIndex + rowCount,
            range.lastRow() + rowCount,
            range.firstColumn(),
            range.lastColumn()));
    return List.copyOf(shifted);
  }

  private static List<String> shiftedRangesForInsertedColumns(
      XSSFSheet sheet, String rawRange, CellBounds range, int columnIndex, int columnCount) {
    if (range.lastColumn() < columnIndex) {
      return List.of(rawRange);
    }
    if (range.firstColumn() >= columnIndex) {
      return List.of(
          shiftedRangeForColumns(
              sheet,
              rawRange,
              range.firstRow(),
              range.lastRow(),
              range.firstColumn() + columnCount,
              range.lastColumn() + columnCount));
    }
    List<String> shifted = new ArrayList<>(2);
    shifted.add(
        SheetStructure.formatRange(
            new CellBounds(
                range.firstRow(), range.lastRow(), range.firstColumn(), columnIndex - 1)));
    shifted.add(
        shiftedRangeForColumns(
            sheet,
            rawRange,
            range.firstRow(),
            range.lastRow(),
            columnIndex + columnCount,
            range.lastColumn() + columnCount));
    return List.copyOf(shifted);
  }

  private static String shiftedRangeForRows(
      XSSFSheet sheet,
      String rawRange,
      int firstRow,
      int lastRow,
      int firstColumn,
      int lastColumn) {
    if (lastRow > Limit.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          "INSERT_ROWS cannot move data validation "
              + rawRange
              + " on sheet '"
              + sheet.getSheetName()
              + "' beyond the maximum row "
              + (Limit.MAX_ROW_INDEX + 1));
    }
    return SheetStructure.formatRange(new CellBounds(firstRow, lastRow, firstColumn, lastColumn));
  }

  private static String shiftedRangeForColumns(
      XSSFSheet sheet,
      String rawRange,
      int firstRow,
      int lastRow,
      int firstColumn,
      int lastColumn) {
    if (lastColumn > Limit.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          "INSERT_COLUMNS cannot move data validation "
              + rawRange
              + " on sheet '"
              + sheet.getSheetName()
              + "' beyond the maximum column "
              + SpreadsheetVersion.EXCEL2007.getLastColumnName());
    }
    return SheetStructure.formatRange(new CellBounds(firstRow, lastRow, firstColumn, lastColumn));
  }

  private static void shiftValidationFormulas(
      CTDataValidation validation,
      XSSFEvaluationWorkbook evaluationWorkbook,
      FormulaShifter formulaShifter,
      int sheetIndex,
      XSSFSheet sheet) {
    String type = validationType(validation);
    if ("list".equals(type)) {
      String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
      if (shouldShiftValidationListFormula(formula1)) {
        validation.setFormula1(
            shiftedFormula(
                formula1,
                FormulaType.DATAVALIDATION_LIST,
                evaluationWorkbook,
                formulaShifter,
                sheetIndex,
                sheet));
      }
      return;
    }
    if (validation.isSetFormula1() && !validation.getFormula1().isBlank()) {
      validation.setFormula1(
          shiftedFormula(
              validation.getFormula1(),
              FormulaType.CELL,
              evaluationWorkbook,
              formulaShifter,
              sheetIndex,
              sheet));
    }
    if (validation.isSetFormula2() && !validation.getFormula2().isBlank()) {
      validation.setFormula2(
          shiftedFormula(
              validation.getFormula2(),
              FormulaType.CELL,
              evaluationWorkbook,
              formulaShifter,
              sheetIndex,
              sheet));
    }
  }

  private static String shiftedFormula(
      String formula,
      FormulaType formulaType,
      XSSFEvaluationWorkbook evaluationWorkbook,
      FormulaShifter formulaShifter,
      int sheetIndex,
      XSSFSheet sheet) {
    try {
      Ptg[] tokens = FormulaParser.parse(formula, evaluationWorkbook, formulaType, sheetIndex);
      if (!formulaShifter.adjustFormula(tokens, sheetIndex)) {
        return formula;
      }
      return FormulaRenderer.toFormulaString(evaluationWorkbook, tokens);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          "Cannot normalize data-validation formula '"
              + formula
              + "' on sheet '"
              + sheet.getSheetName()
              + "' after the structural insert.",
          exception);
    }
  }

  private static String validationType(CTDataValidation validation) {
    return validation.isSetType() ? validation.getType().toString().toLowerCase(Locale.ROOT) : "";
  }

  private static boolean shouldShiftValidationListFormula(String formula1) {
    return !formula1.isBlank() && !isQuotedListLiteral(formula1);
  }

  private static boolean isQuotedListLiteral(String formula) {
    return formula.length() >= 2 && formula.startsWith("\"") && formula.endsWith("\"");
  }
}
