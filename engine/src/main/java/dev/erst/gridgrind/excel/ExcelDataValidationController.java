package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
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
      validation.setErrorStyle(validationDefinition.errorAlert().style().toPoiErrorStyle());
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

    List<ExcelDataValidationSnapshot> snapshots = new ArrayList<>();
    for (XSSFDataValidation validation : sheet.getDataValidations()) {
      List<String> ranges = ranges(validation.getRegions());
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
    return sheet.getDataValidations().size();
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
            findings.add(
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
                    WorkbookAnalysis.AnalysisSeverity.WARNING,
                    "Unsupported data-validation rule",
                    unsupported.detail(),
                    new WorkbookAnalysis.AnalysisLocation.Range(
                        sheetName, unsupported.ranges().getFirst()),
                    unsupported.ranges()));
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
          helper.createExplicitListConstraint(explicitList.values().toArray(String[]::new));
      case ExcelDataValidationRule.FormulaList formulaList ->
          helper.createFormulaListConstraint(formulaList.formula());
      case ExcelDataValidationRule.WholeNumber wholeNumber ->
          helper.createIntegerConstraint(
              wholeNumber.operator().toPoiComparisonOperator(),
              wholeNumber.formula1(),
              wholeNumber.formula2());
      case ExcelDataValidationRule.DecimalNumber decimalNumber ->
          helper.createDecimalConstraint(
              decimalNumber.operator().toPoiComparisonOperator(),
              decimalNumber.formula1(),
              decimalNumber.formula2());
      case ExcelDataValidationRule.DateRule dateRule ->
          helper.createDateConstraint(
              dateRule.operator().toPoiComparisonOperator(),
              dateRule.formula1(),
              dateRule.formula2(),
              null);
      case ExcelDataValidationRule.TimeRule timeRule ->
          helper.createTimeConstraint(
              timeRule.operator().toPoiComparisonOperator(),
              timeRule.formula1(),
              timeRule.formula2());
      case ExcelDataValidationRule.TextLength textLength ->
          helper.createTextLengthConstraint(
              textLength.operator().toPoiComparisonOperator(),
              textLength.formula1(),
              textLength.formula2());
      case ExcelDataValidationRule.CustomFormula customFormula ->
          helper.createCustomConstraint(customFormula.formula());
    };
  }

  private static List<String> ranges(CellRangeAddressList regions) {
    List<String> ranges = new ArrayList<>();
    for (CellRangeAddress region : regions.getCellRangeAddresses()) {
      ranges.add(region.formatAsString());
    }
    return List.copyOf(ranges);
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

  private static ExcelDataValidationSnapshot listSnapshot(
      XSSFDataValidation validation, List<String> ranges) {
    DataValidationConstraint constraint = validation.getValidationConstraint();
    String[] explicitValues = constraint.getExplicitListValues();
    if (explicitValues != null) {
      if (explicitValues.length == 0) {
        return new ExcelDataValidationSnapshot.Unsupported(
            ranges, "INVALID_EXPLICIT_LIST", "Explicit list has no values.");
      }
      return new ExcelDataValidationSnapshot.Supported(
          ranges,
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.ExplicitList(List.of(explicitValues)),
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
      operator = ExcelComparisonOperator.fromPoiComparisonOperator(constraint.getOperator());
    } catch (IllegalArgumentException exception) {
      return new ExcelDataValidationSnapshot.Unsupported(
          ranges,
          "UNKNOWN_OPERATOR",
          family.label + " validation uses an unsupported comparison operator.");
    }
    ExcelDataValidationRule rule =
        switch (family) {
          case INTEGER ->
              new ExcelDataValidationRule.WholeNumber(operator, formula1, constraint.getFormula2());
          case DECIMAL ->
              new ExcelDataValidationRule.DecimalNumber(
                  operator, formula1, constraint.getFormula2());
          case DATE ->
              new ExcelDataValidationRule.DateRule(operator, formula1, constraint.getFormula2());
          case TIME ->
              new ExcelDataValidationRule.TimeRule(operator, formula1, constraint.getFormula2());
          case TEXT_LENGTH ->
              new ExcelDataValidationRule.TextLength(operator, formula1, constraint.getFormula2());
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

  private static ExcelDataValidationPrompt prompt(XSSFDataValidation validation) {
    String title = validation.getPromptBoxTitle();
    String text = validation.getPromptBoxText();
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return null;
    }
    return new ExcelDataValidationPrompt(title, text, validation.getShowPromptBox());
  }

  private static ExcelDataValidationErrorAlert errorAlert(XSSFDataValidation validation) {
    String title = validation.getErrorBoxTitle();
    String text = validation.getErrorBoxText();
    if (title == null || title.isBlank() || text == null || text.isBlank()) {
      return null;
    }
    return new ExcelDataValidationErrorAlert(
        ExcelDataValidationErrorStyle.fromPoiErrorStyle(validation.getErrorStyle()),
        title,
        text,
        validation.getShowErrorBox());
  }

  private static List<WorkbookAnalysis.AnalysisFinding> supportedFindings(
      String sheetName, ExcelDataValidationSnapshot.Supported supported) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    WorkbookAnalysis.AnalysisLocation.Range location =
        new WorkbookAnalysis.AnalysisLocation.Range(sheetName, supported.ranges().getFirst());
    switch (supported.validation().rule()) {
      case ExcelDataValidationRule.ExplicitList _ -> {}
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
