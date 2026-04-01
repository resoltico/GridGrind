package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;

/** Tests for direct data-validation authoring, introspection, and health analysis. */
class ExcelDataValidationControllerTest {
  private final ExcelDataValidationController controller = new ExcelDataValidationController();

  @Test
  void roundTripsExplicitListValidation() throws IOException {
    assertRoundTrip("A1:A3", explicitListValidation());
  }

  @Test
  void roundTripsFormulaListValidation() throws IOException {
    assertRoundTrip(
        "B1:B3",
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.FormulaList("=Statuses"), false, true, null, null));
  }

  @Test
  void roundTripsWholeNumberValidation() throws IOException {
    assertRoundTrip(
        "C1:C3",
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.WholeNumber(
                ExcelComparisonOperator.GREATER_THAN, "=1", null),
            false,
            false,
            null,
            null));
  }

  @Test
  void roundTripsDecimalValidation() throws IOException {
    assertRoundTrip(
        "D1:D3",
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.DecimalNumber(
                ExcelComparisonOperator.GREATER_THAN, "=0.5", null),
            false,
            false,
            null,
            warningAlert()));
  }

  @Test
  void roundTripsDateValidation() throws IOException {
    assertRoundTrip(
        "E1:E3",
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.DateRule(
                ExcelComparisonOperator.EQUAL, "=DATE(2026,4,1)", null),
            false,
            false,
            null,
            null));
  }

  @Test
  void roundTripsTimeValidation() throws IOException {
    assertRoundTrip(
        "F1:F3",
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.TimeRule(
                ExcelComparisonOperator.GREATER_THAN, "=TIME(9,0,0)", null),
            false,
            false,
            null,
            null));
  }

  @Test
  void roundTripsTextLengthValidation() throws IOException {
    assertRoundTrip(
        "G1:G3",
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.TextLength(
                ExcelComparisonOperator.GREATER_THAN, "=20", null),
            false,
            false,
            null,
            null));
  }

  @Test
  void roundTripsCustomFormulaValidation() throws IOException {
    assertRoundTrip(
        "H1:H3",
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.CustomFormula("=LEN(H1)>0"), false, false, null, null));
  }

  @Test
  void selectedSelectionFiltersByIntersection() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, "A1:A3", explicitListValidation());
      controller.setDataValidation(
          sheet,
          "C1:C3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.TextLength(
                  ExcelComparisonOperator.GREATER_THAN, "20", null),
              false,
              false,
              null,
              null));

      List<ExcelDataValidationSnapshot> selected =
          controller.dataValidations(sheet, new ExcelRangeSelection.Selected(List.of("C2")));

      assertEquals(
          List.of(
              new ExcelDataValidationSnapshot.Supported(
                  List.of("C1:C3"),
                  expectedRoundTripDefinition(
                      new ExcelDataValidationDefinition(
                          new ExcelDataValidationRule.TextLength(
                              ExcelComparisonOperator.GREATER_THAN, "20", null),
                          false,
                          false,
                          null,
                          null)))),
          selected);
    }
  }

  @Test
  void clearDataValidationsAllRemovesEveryStructure() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, "A1:A3", explicitListValidation());
      controller.setDataValidation(sheet, "B1:B3", explicitListValidation());

      controller.clearDataValidations(sheet, new ExcelRangeSelection.All());

      assertEquals(0, controller.dataValidationCount(sheet));
      assertFalse(sheet.getCTWorksheet().isSetDataValidations());
    }
  }

  @Test
  void clearDataValidationsAllOnEmptySheetDoesNothing() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.clearDataValidations(sheet, new ExcelRangeSelection.All());

      assertEquals(0, controller.dataValidationCount(sheet));
      assertFalse(sheet.getCTWorksheet().isSetDataValidations());
    }
  }

  @Test
  void selectedClearOnEmptySheetDoesNothing() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.clearDataValidations(sheet, new ExcelRangeSelection.Selected(List.of("A1")));

      assertEquals(0, controller.dataValidationCount(sheet));
      assertFalse(sheet.getCTWorksheet().isSetDataValidations());
    }
  }

  @Test
  void selectedClearRemovesWholeStructuresAndUnsetsTheContainer() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, "A1:A3", explicitListValidation());

      controller.clearDataValidations(sheet, new ExcelRangeSelection.Selected(List.of("A1:A3")));

      assertEquals(0, controller.dataValidationCount(sheet));
      assertFalse(sheet.getCTWorksheet().isSetDataValidations());
    }
  }

  @Test
  void selectedClearSplitsRowRangesAndSkipsBlankSqrefEntries() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, "A1:A10", explicitListValidation());
      sheet
          .getCTWorksheet()
          .getDataValidations()
          .getDataValidationArray(0)
          .setSqref(List.of(" ", "A1:A10"));

      controller.clearDataValidations(sheet, new ExcelRangeSelection.Selected(List.of("A3:A4")));

      assertEquals(
          List.of(
              new ExcelDataValidationSnapshot.Supported(
                  List.of("A1:A2", "A5:A10"),
                  expectedRoundTripDefinition(explicitListValidation()))),
          controller.dataValidations(sheet, new ExcelRangeSelection.All()));
    }
  }

  @Test
  void selectedClearSplitsColumnRangesAndLeavesNonIntersectingStructuresUntouched()
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, "A20:C20", explicitListValidation());
      controller.setDataValidation(sheet, "C1:C3", explicitListValidation());

      controller.clearDataValidations(
          sheet,
          new ExcelRangeSelection.Selected(List.of("B20:B20", "A30:A31", "A1:B3", "D20:D20")));

      assertEquals(
          List.of(
              new ExcelDataValidationSnapshot.Supported(
                  List.of("A20", "C20"), expectedRoundTripDefinition(explicitListValidation())),
              new ExcelDataValidationSnapshot.Supported(
                  List.of("C1:C3"), expectedRoundTripDefinition(explicitListValidation()))),
          controller.dataValidations(sheet, new ExcelRangeSelection.All()));
    }
  }

  @Test
  void dataValidationsExposeUnsupportedStructures() {
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("A1"), "ANY", "Excel 'any value' validation is not modeled by GridGrind."),
        invokeToSnapshot(
            new StubConstraint(DataValidationConstraint.ValidationType.ANY, 0, null, null, null),
            List.of("A1")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("B1"),
            "UNKNOWN_OPERATOR",
            "whole-number validation uses an unsupported comparison operator."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER, -99, null, "1", null),
            List.of("B1")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("B2"), "UNKNOWN", "Unsupported data-validation type: 999"),
        invokeToSnapshot(new StubConstraint(999, 0, null, "1", null), List.of("B2")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("C0"), "INVALID_EXPLICIT_LIST", "Explicit list has no values."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.LIST, 0, new String[0], null, null),
            List.of("C0")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("C1"),
            "MISSING_FORMULA",
            "List validation is missing both explicit values and formula1."),
        invokeToSnapshot(
            new StubConstraint(DataValidationConstraint.ValidationType.LIST, 0, null, null, null),
            List.of("C1")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("C2"),
            "MISSING_FORMULA",
            "List validation is missing both explicit values and formula1."),
        invokeToSnapshot(
            new StubConstraint(DataValidationConstraint.ValidationType.LIST, 0, null, " ", null),
            List.of("C2")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("D1"), "MISSING_FORMULA", "whole-number validation is missing formula1."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER,
                ComparisonOperator.GT,
                null,
                null,
                null),
            List.of("D1")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("D2"), "MISSING_FORMULA", "whole-number validation is missing formula1."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER,
                ComparisonOperator.GT,
                null,
                " ",
                null),
            List.of("D2")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("E1"), "MISSING_FORMULA", "Custom formula validation is missing formula1."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.FORMULA, 0, null, null, null),
            List.of("E1")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("E2"), "MISSING_FORMULA", "Custom formula validation is missing formula1."),
        invokeToSnapshot(
            new StubConstraint(DataValidationConstraint.ValidationType.FORMULA, 0, null, " ", null),
            List.of("E2")));
  }

  @Test
  void supportedSnapshotsDropBlankPromptAndErrorMetadata() {
    StubXssfDataValidation blankPromptTitle =
        new StubXssfDataValidation(
            new StubConstraint(
                DataValidationConstraint.ValidationType.LIST,
                0,
                new String[] {"Queued", "Done"},
                null,
                null));
    blankPromptTitle.setShowPromptBox(true);
    blankPromptTitle.createPromptBox(" ", "Choose one workflow state.");
    blankPromptTitle.setShowErrorBox(true);
    blankPromptTitle.setErrorStyle(DataValidation.ErrorStyle.WARNING);
    blankPromptTitle.createErrorBox("Invalid state", " ");

    ExcelDataValidationSnapshot.Supported blankPromptTitleSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(blankPromptTitle, List.of("A1")));
    assertNull(blankPromptTitleSnapshot.validation().prompt());
    assertNull(blankPromptTitleSnapshot.validation().errorAlert());

    StubXssfDataValidation blankPromptText =
        new StubXssfDataValidation(
            new StubConstraint(
                DataValidationConstraint.ValidationType.LIST,
                0,
                new String[] {"Queued", "Done"},
                null,
                null));
    blankPromptText.setShowPromptBox(true);
    blankPromptText.createPromptBox("Status", " ");
    blankPromptText.setShowErrorBox(true);
    blankPromptText.setErrorStyle(DataValidation.ErrorStyle.INFO);
    blankPromptText.createErrorBox(" ", "Choose an allowed value.");

    ExcelDataValidationSnapshot.Supported blankPromptTextSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(blankPromptText, List.of("B1")));
    assertNull(blankPromptTextSnapshot.validation().prompt());
    assertNull(blankPromptTextSnapshot.validation().errorAlert());
  }

  @Test
  void supportedSnapshotsDropNullTextMetadata() {
    StubXssfDataValidation nullPromptText =
        new StubXssfDataValidation(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.LIST,
                    0,
                    new String[] {"Queued", "Done"},
                    null,
                    null))
            .overridePrompt("Status", null, true)
            .overrideError("Invalid state", null, true, DataValidation.ErrorStyle.STOP);

    ExcelDataValidationSnapshot.Supported snapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(nullPromptText, List.of("C1")));

    assertNull(snapshot.validation().prompt());
    assertNull(snapshot.validation().errorAlert());
  }

  @Test
  void dataValidationHealthFindingsCoverBrokenFormulaFamiliesAndOverlap() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(
          sheet,
          "A1:A3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.FormulaList("#REF!"), false, false, null, null));
      controller.setDataValidation(
          sheet,
          "B1:B3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.WholeNumber(
                  ExcelComparisonOperator.BETWEEN, "#REF!", "#REF!"),
              false,
              false,
              null,
              null));
      controller.setDataValidation(
          sheet,
          "C1:C3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.DecimalNumber(
                  ExcelComparisonOperator.GREATER_THAN, "#REF!", null),
              false,
              false,
              null,
              null));
      controller.setDataValidation(
          sheet,
          "D1:D3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.DateRule(
                  ExcelComparisonOperator.GREATER_THAN, "#REF!", null),
              false,
              false,
              null,
              null));
      controller.setDataValidation(
          sheet,
          "E1:E3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.TimeRule(
                  ExcelComparisonOperator.GREATER_THAN, "#REF!", null),
              false,
              false,
              null,
              null));
      controller.setDataValidation(
          sheet,
          "F1:F3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.TextLength(
                  ExcelComparisonOperator.GREATER_THAN, "#REF!", null),
              false,
              false,
              null,
              null));
      controller.setDataValidation(
          sheet,
          "G1:G3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.CustomFormula("#REF!"), false, false, null, null));
      addHelperValidation(sheet, "H1:H3", helperValidation(sheet, 0, 2, 7, 7));
      addHelperValidation(sheet, "H2:H4", helperValidation(sheet, 1, 3, 7, 7));

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.dataValidationHealthFindings("Budget", sheet);

      assertEquals(9, controller.dataValidationCount(sheet));
      assertEquals(8, findings.stream().filter(this::isBrokenFormulaFinding).count());
      assertEquals(1, findings.stream().filter(this::isOverlapFinding).count());
    }
  }

  @Test
  void dataValidationHealthFindingsIncludeUnsupportedAndHarmlessSupportedRules()
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(
          sheet,
          "A1:A3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.WholeNumber(
                  ExcelComparisonOperator.GREATER_THAN, "1", null),
              false,
              false,
              null,
              null));
      controller.setDataValidation(
          sheet,
          "B1:B3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.FormulaList("Statuses"), false, true, null, null));
      controller.setDataValidation(
          sheet,
          "C1:C3",
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.CustomFormula("LEN(C1)>0"), false, false, null, null));
      addRawValidation(sheet, "D1", STDataValidationType.NONE);

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.dataValidationHealthFindings("Budget", sheet);

      assertEquals(4, controller.dataValidationCount(sheet));
      assertEquals(
          List.of(WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE),
          findings.stream().map(WorkbookAnalysis.AnalysisFinding::code).toList());
    }
  }

  private static ExcelDataValidationDefinition explicitListValidation() {
    return new ExcelDataValidationDefinition(
        new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
        true,
        false,
        new ExcelDataValidationPrompt("Status", "Choose one workflow state.", true),
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.STOP,
            "Invalid status",
            "Use one of the allowed values.",
            true));
  }

  private static ExcelDataValidationErrorAlert warningAlert() {
    return new ExcelDataValidationErrorAlert(
        ExcelDataValidationErrorStyle.WARNING,
        "Outside range",
        "Use a value greater than 0.5.",
        true);
  }

  private ExcelDataValidationSnapshot invokeToSnapshot(
      StubConstraint constraint, List<String> ranges) {
    return ExcelDataValidationController.toSnapshot(new StubXssfDataValidation(constraint), ranges);
  }

  private static ExcelDataValidationDefinition expectedRoundTripDefinition(
      ExcelDataValidationDefinition definition) {
    boolean suppressDropDownArrow = true;
    if (definition.rule() instanceof ExcelDataValidationRule.ExplicitList) {
      suppressDropDownArrow = definition.suppressDropDownArrow();
    }
    return new ExcelDataValidationDefinition(
        definition.rule(),
        definition.allowBlank(),
        suppressDropDownArrow,
        definition.prompt(),
        definition.errorAlert());
  }

  private void assertRoundTrip(String range, ExcelDataValidationDefinition definition)
      throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, range, definition);

      List<ExcelDataValidationSnapshot> snapshots =
          controller.dataValidations(sheet, new ExcelRangeSelection.All());

      assertEquals(
          List.of(
              new ExcelDataValidationSnapshot.Supported(
                  List.of(range), expectedRoundTripDefinition(definition))),
          snapshots);
    }
  }

  private static DataValidation helperValidation(
      XSSFSheet sheet, int firstRow, int lastRow, int firstColumn, int lastColumn) {
    DataValidationHelper helper = sheet.getDataValidationHelper();
    return helper.createValidation(
        helper.createExplicitListConstraint(new String[] {"Queued", "Done"}),
        new CellRangeAddressList(firstRow, lastRow, firstColumn, lastColumn));
  }

  private static void addHelperValidation(
      XSSFSheet sheet, String range, DataValidation validation) {
    sheet.addValidationData(validation);
    sheet
        .getCTWorksheet()
        .getDataValidations()
        .getDataValidationArray(
            sheet.getCTWorksheet().getDataValidations().sizeOfDataValidationArray() - 1)
        .setSqref(List.of(range));
    syncCount(sheet);
  }

  private static void syncCount(XSSFSheet sheet) {
    sheet
        .getCTWorksheet()
        .getDataValidations()
        .setCount(sheet.getCTWorksheet().getDataValidations().sizeOfDataValidationArray());
  }

  private static void addRawValidation(
      XSSFSheet sheet, String range, STDataValidationType.Enum type) {
    CTDataValidations dataValidations =
        sheet.getCTWorksheet().isSetDataValidations()
            ? sheet.getCTWorksheet().getDataValidations()
            : sheet.getCTWorksheet().addNewDataValidations();
    CTDataValidation validation = dataValidations.addNewDataValidation();
    validation.setType(type);
    validation.setSqref(List.of(range));
    syncCount(sheet);
  }

  /** Immutable fake POI constraint used to drive controller snapshot branches deterministically. */
  private static final class StubConstraint implements DataValidationConstraint {
    private final int validationType;
    private final int operator;
    private final String[] explicitListValues;
    private final String formula1;
    private final String formula2;

    private StubConstraint(
        int validationType,
        int operator,
        String[] explicitListValues,
        String formula1,
        String formula2) {
      this.validationType = validationType;
      this.operator = operator;
      this.explicitListValues = explicitListValues;
      this.formula1 = formula1;
      this.formula2 = formula2;
    }

    @Override
    public int getValidationType() {
      return validationType;
    }

    @Override
    public int getOperator() {
      return operator;
    }

    @Override
    public void setOperator(int ignored) {}

    @Override
    public String[] getExplicitListValues() {
      return explicitListValues == null ? null : explicitListValues.clone();
    }

    @Override
    public void setExplicitListValues(String[] ignored) {}

    @Override
    public String getFormula1() {
      return formula1;
    }

    @Override
    public void setFormula1(String ignored) {}

    @Override
    public String getFormula2() {
      return formula2;
    }

    @Override
    public void setFormula2(String ignored) {}
  }

  /** Minimal XSSF wrapper whose constraint can be supplied directly by the test. */
  private static final class StubXssfDataValidation extends XSSFDataValidation {
    private final DataValidationConstraint constraint;
    private boolean promptOverridden;
    private String promptTitle;
    private String promptText;
    private boolean showPromptBox;
    private boolean errorOverridden;
    private String errorTitle;
    private String errorText;
    private boolean showErrorBox;
    private int errorStyle;

    private StubXssfDataValidation(DataValidationConstraint constraint) {
      super(
          new XSSFDataValidationConstraint(new String[] {"Queued"}),
          new CellRangeAddressList(0, 0, 0, 0),
          CTDataValidation.Factory.newInstance());
      this.constraint = constraint;
    }

    private StubXssfDataValidation overridePrompt(String title, String text, boolean show) {
      promptOverridden = true;
      promptTitle = title;
      promptText = text;
      showPromptBox = show;
      return this;
    }

    private StubXssfDataValidation overrideError(
        String title, String text, boolean show, int style) {
      errorOverridden = true;
      errorTitle = title;
      errorText = text;
      showErrorBox = show;
      errorStyle = style;
      return this;
    }

    @Override
    public DataValidationConstraint getValidationConstraint() {
      return constraint;
    }

    @Override
    public String getPromptBoxTitle() {
      return promptOverridden ? promptTitle : super.getPromptBoxTitle();
    }

    @Override
    public String getPromptBoxText() {
      return promptOverridden ? promptText : super.getPromptBoxText();
    }

    @Override
    public boolean getShowPromptBox() {
      return promptOverridden ? showPromptBox : super.getShowPromptBox();
    }

    @Override
    public String getErrorBoxTitle() {
      return errorOverridden ? errorTitle : super.getErrorBoxTitle();
    }

    @Override
    public String getErrorBoxText() {
      return errorOverridden ? errorText : super.getErrorBoxText();
    }

    @Override
    public boolean getShowErrorBox() {
      return errorOverridden ? showErrorBox : super.getShowErrorBox();
    }

    @Override
    public int getErrorStyle() {
      return errorOverridden ? errorStyle : super.getErrorStyle();
    }
  }

  private boolean isBrokenFormulaFinding(WorkbookAnalysis.AnalysisFinding finding) {
    return finding.code() == WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA;
  }

  private boolean isOverlapFinding(WorkbookAnalysis.AnalysisFinding finding) {
    return finding.code() == WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_OVERLAPPING_RULES;
  }
}
