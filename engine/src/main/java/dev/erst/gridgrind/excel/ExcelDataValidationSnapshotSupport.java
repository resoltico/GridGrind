package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import javax.xml.namespace.QName;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Reads workbook data-validation XML into canonical GridGrind snapshot values. */
final class ExcelDataValidationSnapshotSupport {
  List<ExcelDataValidationSnapshot> dataValidations(
      XSSFSheet sheet, ExcelRangeSelection selection) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    if (dataValidations == null) {
      return List.of();
    }
    List<ExcelDataValidationSnapshot> snapshots = new ArrayList<>();
    for (CTDataValidation validation : dataValidations.getDataValidationArray()) {
      List<String> ranges = ExcelSqrefSupport.normalizedSqref(validation.getSqref());
      if (!ExcelDataValidationRangeSupport.matchesSelection(ranges, selection)) {
        continue;
      }
      snapshots.add(toSnapshot(validation, ranges));
    }
    return List.copyOf(snapshots);
  }

  static ExcelDataValidationSnapshot toSnapshot(
      XSSFDataValidation validation, List<String> ranges) {
    DataValidationConstraint constraint = validation.getValidationConstraint();
    return switch (constraint.getValidationType()) {
      case DataValidationConstraint.ValidationType.ANY ->
          new ExcelDataValidationSnapshot.Unsupported(
              ranges, "ANY", "Excel 'any value' validation is not modeled by GridGrind.");
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
          new ExcelDataValidationSnapshot.Unsupported(
              ranges,
              "UNKNOWN",
              "Unsupported data-validation type: " + constraint.getValidationType());
    };
  }

  static ExcelDataValidationSnapshot toSnapshot(CTDataValidation validation, List<String> ranges) {
    String type = Objects.requireNonNullElse(rawAttribute(validation, "type"), "");
    return switch (type.toLowerCase(Locale.ROOT)) {
      case "none", "" ->
          new ExcelDataValidationSnapshot.Unsupported(
              ranges, "ANY", "Excel 'any value' validation is not modeled by GridGrind.");
      case "list" -> listSnapshot(validation, ranges);
      case "whole" -> comparisonSnapshot(validation, ranges, ComparisonFamily.INTEGER);
      case "decimal" -> comparisonSnapshot(validation, ranges, ComparisonFamily.DECIMAL);
      case "date" -> comparisonSnapshot(validation, ranges, ComparisonFamily.DATE);
      case "time" -> comparisonSnapshot(validation, ranges, ComparisonFamily.TIME);
      case "textlength" -> comparisonSnapshot(validation, ranges, ComparisonFamily.TEXT_LENGTH);
      case "custom" -> customFormulaSnapshot(validation, ranges);
      default ->
          new ExcelDataValidationSnapshot.Unsupported(
              ranges,
              "UNKNOWN",
              "Unsupported data-validation type: " + Objects.requireNonNullElse(type, "UNKNOWN"));
    };
  }

  private static ExcelDataValidationSnapshot listSnapshot(
      XSSFDataValidation validation, List<String> ranges) {
    DataValidationConstraint constraint = validation.getValidationConstraint();
    String[] explicitValues = constraint.getExplicitListValues();
    if (explicitValues != null) {
      List<String> explicitListValues =
          "\"\"".equals(Objects.requireNonNullElse(constraint.getFormula1(), ""))
              ? List.of()
              : List.of(explicitValues);
      return new ExcelDataValidationSnapshot.Supported(
          ranges,
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.ExplicitList(explicitListValues),
              validation.getEmptyCellAllowed(),
              validation.getSuppressDropDownArrow(),
              prompt(validation),
              errorAlert(validation)));
    }
    String formula1 = constraint.getFormula1();
    if (formula1 == null || formula1.isBlank()) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges,
          "MISSING_FORMULA",
          "List validation is missing both explicit values and formula1.");
    }
    return new ExcelDataValidationSnapshot.Supported(
        ranges,
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.FormulaList(formula1),
            validation.getEmptyCellAllowed(),
            validation.getSuppressDropDownArrow(),
            prompt(validation),
            errorAlert(validation)));
  }

  private static ExcelDataValidationSnapshot listSnapshot(
      CTDataValidation validation, List<String> ranges) {
    String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
    if (formula1.isBlank()) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges,
          "MISSING_FORMULA",
          "List validation is missing both explicit values and formula1.");
    }
    ExcelDataValidationRule rule;
    if ("\"\"".equals(formula1)) {
      rule = new ExcelDataValidationRule.ExplicitList(List.of());
    } else if (formula1.length() >= 2 && formula1.startsWith("\"") && formula1.endsWith("\"")) {
      String rawValues = formula1.substring(1, formula1.length() - 1);
      rule = new ExcelDataValidationRule.ExplicitList(Arrays.asList(rawValues.split(",")));
    } else {
      rule = new ExcelDataValidationRule.FormulaList(formula1);
    }
    return new ExcelDataValidationSnapshot.Supported(
        ranges,
        new ExcelDataValidationDefinition(
            rule,
            validation.isSetAllowBlank() && validation.getAllowBlank(),
            suppressDropDownArrow(validation),
            prompt(validation),
            errorAlert(validation)));
  }

  private static ExcelDataValidationSnapshot comparisonSnapshot(
      XSSFDataValidation validation,
      List<String> ranges,
      DataValidationConstraint constraint,
      ComparisonFamily family) {
    String formula1 = constraint.getFormula1();
    if (formula1 == null || formula1.isBlank()) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges, "MISSING_FORMULA", family.label + " validation is missing formula1.");
    }
    ExcelComparisonOperator operator;
    try {
      operator = ExcelComparisonOperatorPoiBridge.fromPoi(constraint.getOperator());
    } catch (IllegalArgumentException exception) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges,
          "UNKNOWN_OPERATOR",
          family.label + " validation uses an unsupported comparison operator.");
    }
    String formula2 = constraint.getFormula2();
    if (requiresUpperBound(operator) && (formula2 == null || formula2.isBlank())) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges,
          "MISSING_FORMULA",
          family.label + " validation is missing formula2 for " + operatorLabel(operator) + ".");
    }
    ExcelDataValidationRule rule = comparisonRule(family, operator, formula1, formula2);
    return new ExcelDataValidationSnapshot.Supported(
        ranges,
        new ExcelDataValidationDefinition(
            rule,
            validation.getEmptyCellAllowed(),
            validation.getSuppressDropDownArrow(),
            prompt(validation),
            errorAlert(validation)));
  }

  private static ExcelDataValidationSnapshot comparisonSnapshot(
      CTDataValidation validation, List<String> ranges, ComparisonFamily family) {
    String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
    if (formula1.isBlank()) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges, "MISSING_FORMULA", family.label + " validation is missing formula1.");
    }
    ExcelComparisonOperator operator;
    try {
      operator = comparisonOperator(validation);
    } catch (IllegalArgumentException exception) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges,
          "UNKNOWN_OPERATOR",
          family.label + " validation uses an unsupported comparison operator.");
    }
    String formula2 = validation.isSetFormula2() ? validation.getFormula2() : null;
    if (requiresUpperBound(operator) && (formula2 == null || formula2.isBlank())) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges,
          "MISSING_FORMULA",
          family.label + " validation is missing formula2 for " + operatorLabel(operator) + ".");
    }
    ExcelDataValidationRule rule = comparisonRule(family, operator, formula1, formula2);
    return new ExcelDataValidationSnapshot.Supported(
        ranges,
        new ExcelDataValidationDefinition(
            rule,
            validation.isSetAllowBlank() && validation.getAllowBlank(),
            suppressDropDownArrow(validation),
            prompt(validation),
            errorAlert(validation)));
  }

  private static ExcelDataValidationSnapshot customFormulaSnapshot(
      XSSFDataValidation validation, List<String> ranges) {
    String formula1 = validation.getValidationConstraint().getFormula1();
    if (formula1 == null || formula1.isBlank()) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges, "MISSING_FORMULA", "Custom formula validation is missing formula1.");
    }
    return new ExcelDataValidationSnapshot.Supported(
        ranges,
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.CustomFormula(formula1),
            validation.getEmptyCellAllowed(),
            validation.getSuppressDropDownArrow(),
            prompt(validation),
            errorAlert(validation)));
  }

  private static ExcelDataValidationSnapshot customFormulaSnapshot(
      CTDataValidation validation, List<String> ranges) {
    String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
    if (formula1.isBlank()) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges, "MISSING_FORMULA", "Custom formula validation is missing formula1.");
    }
    return new ExcelDataValidationSnapshot.Supported(
        ranges,
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.CustomFormula(formula1),
            validation.isSetAllowBlank() && validation.getAllowBlank(),
            suppressDropDownArrow(validation),
            prompt(validation),
            errorAlert(validation)));
  }

  static ExcelDataValidationPrompt prompt(XSSFDataValidation validation) {
    String title = validation.getPromptBoxTitle();
    String text = validation.getPromptBoxText();
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return null;
    }
    return new ExcelDataValidationPrompt(title, text, validation.getShowPromptBox());
  }

  static ExcelDataValidationPrompt prompt(CTDataValidation validation) {
    String title = validation.isSetPromptTitle() ? validation.getPromptTitle() : null;
    String text = validation.isSetPrompt() ? validation.getPrompt() : null;
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return null;
    }
    return new ExcelDataValidationPrompt(
        title, text, validation.isSetShowInputMessage() && validation.getShowInputMessage());
  }

  static ExcelDataValidationErrorAlert errorAlert(XSSFDataValidation validation) {
    String title = validation.getErrorBoxTitle();
    String text = validation.getErrorBoxText();
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return null;
    }
    return new ExcelDataValidationErrorAlert(
        ExcelDataValidationPoiBridge.fromPoiErrorStyle(validation.getErrorStyle()),
        title,
        text,
        validation.getShowErrorBox());
  }

  static ExcelDataValidationErrorAlert errorAlert(CTDataValidation validation) {
    String title = validation.isSetErrorTitle() ? validation.getErrorTitle() : null;
    String text = validation.isSetError() ? validation.getError() : null;
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return null;
    }
    return new ExcelDataValidationErrorAlert(
        errorStyle(validation),
        title,
        text,
        validation.isSetShowErrorMessage() && validation.getShowErrorMessage());
  }

  static ExcelComparisonOperator comparisonOperator(CTDataValidation validation) {
    String operator = Objects.requireNonNullElse(rawAttribute(validation, "operator"), "between");
    return switch (operator) {
      case "between" -> ExcelComparisonOperator.BETWEEN;
      case "notBetween" -> ExcelComparisonOperator.NOT_BETWEEN;
      case "equal" -> ExcelComparisonOperator.NOT_EQUAL;
      case "notEqual" -> ExcelComparisonOperator.EQUAL;
      case "greaterThan" -> ExcelComparisonOperator.LESS_THAN;
      case "lessThan" -> ExcelComparisonOperator.GREATER_THAN;
      case "greaterThanOrEqual" -> ExcelComparisonOperator.LESS_OR_EQUAL;
      case "lessThanOrEqual" -> ExcelComparisonOperator.GREATER_OR_EQUAL;
      default ->
          throw new IllegalArgumentException("Unsupported raw validation operator: " + operator);
    };
  }

  static ExcelDataValidationErrorStyle errorStyle(CTDataValidation validation) {
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

  static String operatorLabel(ExcelComparisonOperator operator) {
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

  private static ExcelDataValidationRule comparisonRule(
      ComparisonFamily family, ExcelComparisonOperator operator, String formula1, String formula2) {
    return switch (family) {
      case INTEGER -> new ExcelDataValidationRule.WholeNumber(operator, formula1, formula2);
      case DECIMAL -> new ExcelDataValidationRule.DecimalNumber(operator, formula1, formula2);
      case DATE -> new ExcelDataValidationRule.DateRule(operator, formula1, formula2);
      case TIME -> new ExcelDataValidationRule.TimeRule(operator, formula1, formula2);
      case TEXT_LENGTH -> new ExcelDataValidationRule.TextLength(operator, formula1, formula2);
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
}
