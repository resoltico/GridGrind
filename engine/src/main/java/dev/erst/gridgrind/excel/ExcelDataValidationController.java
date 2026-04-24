package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import javax.xml.namespace.QName;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Reads, writes, and analyzes Excel data-validation structures on one XSSF sheet. */
final class ExcelDataValidationController {
  /** Creates or replaces one data-validation rule over the requested sheet range. */
  void setDataValidation(
      XSSFSheet sheet, String range, ExcelDataValidationDefinition validationDefinition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(validationDefinition, "validationDefinition must not be null");

    ExcelRange targetRange = ExcelRange.parse(range);
    normalizeOverlappingSqref(sheet, List.of(targetRange));

    DataValidationHelper helper = sheet.getDataValidationHelper();
    DataValidationConstraint constraint = toConstraint(helper, validationDefinition.rule());
    CellRangeAddressList regions =
        new CellRangeAddressList(
            targetRange.firstRow(),
            targetRange.lastRow(),
            targetRange.firstColumn(),
            targetRange.lastColumn());
    DataValidation validation = helper.createValidation(constraint, regions);
    validation.setEmptyCellAllowed(validationDefinition.allowBlank());
    validation.setSuppressDropDownArrow(validationDefinition.suppressDropDownArrow());
    if (validationDefinition.prompt() == null) {
      validation.setShowPromptBox(false);
    } else {
      validation.setShowPromptBox(validationDefinition.prompt().showPromptBox());
      validation.createPromptBox(
          validationDefinition.prompt().title(), validationDefinition.prompt().text());
    }
    if (validationDefinition.errorAlert() == null) {
      validation.setShowErrorBox(false);
    } else {
      validation.setShowErrorBox(validationDefinition.errorAlert().showErrorBox());
      validation.setErrorStyle(
          ExcelDataValidationPoiBridge.toPoiErrorStyle(validationDefinition.errorAlert().style()));
      validation.createErrorBox(
          validationDefinition.errorAlert().title(), validationDefinition.errorAlert().text());
    }
    sheet.addValidationData(validation);
    syncValidationCount(sheet);
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  void clearDataValidations(XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    switch (selection) {
      case ExcelRangeSelection.All _ -> {
        if (sheet.getCTWorksheet().isSetDataValidations()) {
          sheet.getCTWorksheet().unsetDataValidations();
        }
      }
      case ExcelRangeSelection.Selected selected ->
          normalizeOverlappingSqref(
              sheet, selected.ranges().stream().map(ExcelRange::parse).toList());
    }
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  List<ExcelDataValidationSnapshot> dataValidations(
      XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    if (dataValidations == null) {
      return List.of();
    }
    List<ExcelDataValidationSnapshot> snapshots = new ArrayList<>();
    for (CTDataValidation validation : dataValidations.getDataValidationArray()) {
      List<String> ranges = ExcelSqrefSupport.normalizedSqref(validation.getSqref());
      if (!matchesSelection(ranges, selection)) {
        continue;
      }
      snapshots.add(toSnapshot(validation, ranges));
    }
    return List.copyOf(snapshots);
  }

  /** Returns the number of raw data-validation structures currently present on the sheet. */
  int dataValidationCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    return dataValidations == null ? 0 : dataValidations.sizeOfDataValidationArray();
  }

  /** Returns derived findings for data-validation health on this sheet. */
  List<WorkbookAnalysis.AnalysisFinding> dataValidationHealthFindings(
      String sheetName, XSSFSheet sheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");

    List<ExcelDataValidationSnapshot> snapshots =
        dataValidations(sheet, new ExcelRangeSelection.All());
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

  private void normalizeOverlappingSqref(XSSFSheet sheet, List<ExcelRange> cutouts) {
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

  private List<String> retainedRanges(CTDataValidation validation, List<ExcelRange> cutouts) {
    List<ExcelRange> retained =
        ExcelSqrefSupport.normalizedSqref(validation.getSqref()).stream()
            .map(ExcelRange::parse)
            .toList();
    List<ExcelRange> next = new ArrayList<>();
    for (ExcelRange cutout : cutouts) {
      next.clear();
      for (ExcelRange existingRange : retained) {
        next.addAll(subtract(existingRange, cutout));
      }
      retained = List.copyOf(next);
    }
    return retained.stream().map(ExcelDataValidationController::formatRange).toList();
  }

  private static List<ExcelRange> subtract(ExcelRange source, ExcelRange cutout) {
    if (!intersects(source, cutout)) {
      return List.of(source);
    }
    int intersectionFirstRow = Math.max(source.firstRow(), cutout.firstRow());
    int intersectionLastRow = Math.min(source.lastRow(), cutout.lastRow());
    int intersectionFirstColumn = Math.max(source.firstColumn(), cutout.firstColumn());
    int intersectionLastColumn = Math.min(source.lastColumn(), cutout.lastColumn());

    List<ExcelRange> retained = new ArrayList<>(4);
    if (source.firstRow() < intersectionFirstRow) {
      retained.add(
          new ExcelRange(
              source.firstRow(),
              intersectionFirstRow - 1,
              source.firstColumn(),
              source.lastColumn()));
    }
    if (intersectionLastRow < source.lastRow()) {
      retained.add(
          new ExcelRange(
              intersectionLastRow + 1,
              source.lastRow(),
              source.firstColumn(),
              source.lastColumn()));
    }
    if (source.firstColumn() < intersectionFirstColumn) {
      retained.add(
          new ExcelRange(
              intersectionFirstRow,
              intersectionLastRow,
              source.firstColumn(),
              intersectionFirstColumn - 1));
    }
    if (intersectionLastColumn < source.lastColumn()) {
      retained.add(
          new ExcelRange(
              intersectionFirstRow,
              intersectionLastRow,
              intersectionLastColumn + 1,
              source.lastColumn()));
    }
    return List.copyOf(retained);
  }

  private static boolean intersects(ExcelRange first, ExcelRange second) {
    return first.firstRow() <= second.lastRow()
        && first.lastRow() >= second.firstRow()
        && first.firstColumn() <= second.lastColumn()
        && first.lastColumn() >= second.firstColumn();
  }

  private static String formatRange(ExcelRange range) {
    return new CellRangeAddress(
            range.firstRow(), range.lastRow(), range.firstColumn(), range.lastColumn())
        .formatAsString();
  }

  private static void syncValidationCount(XSSFSheet sheet) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    int count = dataValidations.sizeOfDataValidationArray();
    if (count == 0) {
      sheet.getCTWorksheet().unsetDataValidations();
      return;
    }
    dataValidations.setCount(count);
  }

  private static DataValidationConstraint toConstraint(
      DataValidationHelper helper, ExcelDataValidationRule rule) {
    return switch (rule) {
      case ExcelDataValidationRule.ExplicitList explicitList ->
          explicitList.values().isEmpty()
              ? helper.createFormulaListConstraint("\"\"")
              : helper.createExplicitListConstraint(explicitList.values().toArray(String[]::new));
      case ExcelDataValidationRule.FormulaList formulaList ->
          helper.createFormulaListConstraint(formulaList.formula());
      case ExcelDataValidationRule.WholeNumber wholeNumber ->
          helper.createIntegerConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(wholeNumber.operator()),
              wholeNumber.formula1(),
              wholeNumber.formula2());
      case ExcelDataValidationRule.DecimalNumber decimalNumber ->
          helper.createDecimalConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(decimalNumber.operator()),
              decimalNumber.formula1(),
              decimalNumber.formula2());
      case ExcelDataValidationRule.DateRule dateRule ->
          helper.createDateConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(dateRule.operator()),
              dateRule.formula1(),
              dateRule.formula2(),
              null);
      case ExcelDataValidationRule.TimeRule timeRule ->
          helper.createTimeConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(timeRule.operator()),
              timeRule.formula1(),
              timeRule.formula2());
      case ExcelDataValidationRule.TextLength textLength ->
          helper.createTextLengthConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(textLength.operator()),
              textLength.formula1(),
              textLength.formula2());
      case ExcelDataValidationRule.CustomFormula customFormula ->
          helper.createCustomConstraint(customFormula.formula());
    };
  }

  private static boolean matchesSelection(List<String> ranges, ExcelRangeSelection selection) {
    return switch (selection) {
      case ExcelRangeSelection.All _ -> true;
      case ExcelRangeSelection.Selected selected ->
          ranges.stream()
              .map(ExcelRange::parse)
              .anyMatch(
                  existing ->
                      selected.ranges().stream()
                          .map(ExcelRange::parse)
                          .anyMatch(selectedRange -> intersects(existing, selectedRange)));
    };
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
    Objects.requireNonNull(validation, "validation must not be null");
    Objects.requireNonNull(ranges, "ranges must not be null");

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
    ExcelDataValidationRule rule =
        switch (family) {
          case INTEGER -> new ExcelDataValidationRule.WholeNumber(operator, formula1, formula2);
          case DECIMAL -> new ExcelDataValidationRule.DecimalNumber(operator, formula1, formula2);
          case DATE -> new ExcelDataValidationRule.DateRule(operator, formula1, formula2);
          case TIME -> new ExcelDataValidationRule.TimeRule(operator, formula1, formula2);
          case TEXT_LENGTH -> new ExcelDataValidationRule.TextLength(operator, formula1, formula2);
        };
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
    ExcelDataValidationRule rule =
        switch (family) {
          case INTEGER -> new ExcelDataValidationRule.WholeNumber(operator, formula1, formula2);
          case DECIMAL -> new ExcelDataValidationRule.DecimalNumber(operator, formula1, formula2);
          case DATE -> new ExcelDataValidationRule.DateRule(operator, formula1, formula2);
          case TIME -> new ExcelDataValidationRule.TimeRule(operator, formula1, formula2);
          case TEXT_LENGTH -> new ExcelDataValidationRule.TextLength(operator, formula1, formula2);
        };
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
      return java.util.Optional.<ExcelDataValidationPrompt>empty().orElse(null);
    }
    return new ExcelDataValidationPrompt(title, text, validation.getShowPromptBox());
  }

  static ExcelDataValidationPrompt prompt(CTDataValidation validation) {
    String title = validation.isSetPromptTitle() ? validation.getPromptTitle() : null;
    String text = validation.isSetPrompt() ? validation.getPrompt() : null;
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return java.util.Optional.<ExcelDataValidationPrompt>empty().orElse(null);
    }
    return new ExcelDataValidationPrompt(
        title, text, validation.isSetShowInputMessage() && validation.getShowInputMessage());
  }

  static ExcelDataValidationErrorAlert errorAlert(XSSFDataValidation validation) {
    String title = validation.getErrorBoxTitle();
    String text = validation.getErrorBoxText();
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return java.util.Optional.<ExcelDataValidationErrorAlert>empty().orElse(null);
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
      return java.util.Optional.<ExcelDataValidationErrorAlert>empty().orElse(null);
    }
    return new ExcelDataValidationErrorAlert(
        errorStyle(validation),
        title,
        text,
        validation.isSetShowErrorMessage() && validation.getShowErrorMessage());
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
                  WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_EMPTY_EXPLICIT_LIST,
                  WorkbookAnalysis.AnalysisSeverity.ERROR,
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
          WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_MALFORMED_RULE,
          WorkbookAnalysis.AnalysisSeverity.ERROR,
          "Data-validation rule is malformed",
          unsupported.detail(),
          location,
          unsupported.ranges());
    }
    return new WorkbookAnalysis.AnalysisFinding(
        WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
        WorkbookAnalysis.AnalysisSeverity.WARNING,
        "Unsupported data-validation rule",
        unsupported.detail(),
        location,
        unsupported.ranges());
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

  private static boolean suppressDropDownArrow(CTDataValidation validation) {
    return !validation.isSetShowDropDown() || !validation.getShowDropDown();
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

  private static String rawAttribute(CTDataValidation validation, String attributeName) {
    try (var cursor = validation.newCursor()) {
      return cursor.getAttributeText(new QName("", attributeName));
    }
  }

  private static boolean requiresUpperBound(ExcelComparisonOperator operator) {
    return operator == ExcelComparisonOperator.BETWEEN
        || operator == ExcelComparisonOperator.NOT_BETWEEN;
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
              WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
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
            if (!intersects(ExcelRange.parse(firstRange), ExcelRange.parse(secondRange))) {
              continue;
            }
            findings.add(
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_OVERLAPPING_RULES,
                    WorkbookAnalysis.AnalysisSeverity.WARNING,
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

  /** Groups POI's comparison-validation families so finding messages stay precise. */
  private enum ComparisonFamily {
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
