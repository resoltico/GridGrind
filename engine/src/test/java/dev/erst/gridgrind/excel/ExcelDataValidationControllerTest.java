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
            List.of("C0"), "EMPTY_EXPLICIT_LIST", "Explicit list has no values."),
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
  void xssfMetadataHelpersCoverPromptErrorAndSupportedFormulaPayloads() {
    assertNull(
        ExcelDataValidationController.prompt(
            xssfListStub().overridePrompt(null, "Choose one workflow state.", true)));
    assertNull(
        ExcelDataValidationController.prompt(
            xssfListStub().overridePrompt(" ", "Choose one workflow state.", true)));
    assertNull(
        ExcelDataValidationController.prompt(xssfListStub().overridePrompt("Status", null, true)));
    assertNull(
        ExcelDataValidationController.prompt(xssfListStub().overridePrompt("Status", " ", true)));
    assertEquals(
        new ExcelDataValidationPrompt("Status", "Choose one workflow state.", false),
        ExcelDataValidationController.prompt(
            xssfListStub().overridePrompt("Status", "Choose one workflow state.", false)));

    assertNull(
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError(
                    null, "Use an allowed value.", true, DataValidation.ErrorStyle.STOP)));
    assertNull(
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError(
                    " ", "Use an allowed value.", true, DataValidation.ErrorStyle.STOP)));
    assertNull(
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError("Invalid status", null, true, DataValidation.ErrorStyle.STOP)));
    assertNull(
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError("Invalid status", " ", true, DataValidation.ErrorStyle.STOP)));
    assertEquals(
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.INFORMATION, "Heads up", "Use an allowed value.", false),
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError(
                    "Heads up", "Use an allowed value.", false, DataValidation.ErrorStyle.INFO)));

    ExcelDataValidationSnapshot.Supported formulaListSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(
                new StubXssfDataValidation(
                        new StubConstraint(
                            DataValidationConstraint.ValidationType.LIST,
                            0,
                            null,
                            "Statuses",
                            null))
                    .overridePrompt("Status", "Choose one workflow state.", false)
                    .overrideError(
                        "Heads up",
                        "Read the allowed values.",
                        false,
                        DataValidation.ErrorStyle.WARNING),
                List.of("C1:C3")));
    assertEquals(
        new ExcelDataValidationRule.FormulaList("Statuses"),
        formulaListSnapshot.validation().rule());
    assertEquals(
        new ExcelDataValidationPrompt("Status", "Choose one workflow state.", false),
        formulaListSnapshot.validation().prompt());
    assertEquals(
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.WARNING, "Heads up", "Read the allowed values.", false),
        formulaListSnapshot.validation().errorAlert());

    ExcelDataValidationSnapshot.Supported customFormulaSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(
                new StubXssfDataValidation(
                        new StubConstraint(
                            DataValidationConstraint.ValidationType.FORMULA,
                            0,
                            null,
                            "LEN(D1)>0",
                            null))
                    .overridePrompt("Prompt", "Keep the value.", true)
                    .overrideError(
                        "Invalid value",
                        "Formula must pass.",
                        true,
                        DataValidation.ErrorStyle.STOP),
                List.of("D1")));
    assertEquals(
        new ExcelDataValidationRule.CustomFormula("LEN(D1)>0"),
        customFormulaSnapshot.validation().rule());
  }

  @Test
  void rawCtMetadataHelpersCoverPromptAndErrorBranches() {
    CTDataValidation promptDefaults = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    promptDefaults.setPromptTitle("Status");
    promptDefaults.setPrompt("Choose one workflow state.");
    assertEquals(
        new ExcelDataValidationPrompt("Status", "Choose one workflow state.", false),
        ExcelDataValidationController.prompt(promptDefaults));

    CTDataValidation promptBlankTitle =
        rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    promptBlankTitle.setPromptTitle(" ");
    promptBlankTitle.setPrompt("Choose one workflow state.");
    assertNull(ExcelDataValidationController.prompt(promptBlankTitle));

    CTDataValidation promptMissingText =
        rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    promptMissingText.setPromptTitle("Status");
    assertNull(ExcelDataValidationController.prompt(promptMissingText));

    CTDataValidation errorDefaults = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    errorDefaults.setErrorTitle("Invalid status");
    errorDefaults.setError("Use one of the allowed values.");
    assertEquals(
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.STOP,
            "Invalid status",
            "Use one of the allowed values.",
            false),
        ExcelDataValidationController.errorAlert(errorDefaults));

    CTDataValidation warningError = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    warningError.setErrorTitle("Warning");
    warningError.setError("Watch this value.");
    warningError.setShowErrorMessage(true);
    warningError.setErrorStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationErrorStyle.WARNING);
    assertEquals(
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.WARNING, "Warning", "Watch this value.", true),
        ExcelDataValidationController.errorAlert(warningError));

    CTDataValidation infoError = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    infoError.setErrorTitle("Info");
    infoError.setError("Read the list.");
    infoError.setErrorStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationErrorStyle.INFORMATION);
    assertEquals(
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.INFORMATION, "Info", "Read the list.", false),
        ExcelDataValidationController.errorAlert(infoError));

    CTDataValidation blankErrorTitle = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    blankErrorTitle.setErrorTitle(" ");
    blankErrorTitle.setError("Use one of the allowed values.");
    assertNull(ExcelDataValidationController.errorAlert(blankErrorTitle));

    CTDataValidation missingErrorText =
        rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    missingErrorText.setErrorTitle("Invalid status");
    assertNull(ExcelDataValidationController.errorAlert(missingErrorText));
  }

  @Test
  void rawEnumMappingsAndLabelsCoverDefaultedAndMalformedMetadata() {
    CTDataValidation defaultWhole = rawValidation(STDataValidationType.WHOLE, "B1", "1");
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("B1"),
            "MISSING_FORMULA",
            "whole-number validation is missing formula2 for between operator."),
        ExcelDataValidationController.toSnapshot(defaultWhole, List.of("B1")));

    assertEquals(
        ExcelComparisonOperator.BETWEEN,
        ExcelDataValidationController.comparisonOperator(rawWholeOperator(null)));
    assertEquals(
        ExcelComparisonOperator.NOT_BETWEEN,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .NOT_BETWEEN)));
    assertEquals(
        ExcelComparisonOperator.NOT_EQUAL,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .EQUAL)));
    assertEquals(
        ExcelComparisonOperator.EQUAL,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .NOT_EQUAL)));
    assertEquals(
        ExcelComparisonOperator.LESS_THAN,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .GREATER_THAN)));
    assertEquals(
        ExcelComparisonOperator.GREATER_THAN,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .LESS_THAN)));
    assertEquals(
        ExcelComparisonOperator.LESS_OR_EQUAL,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .GREATER_THAN_OR_EQUAL)));
    assertEquals(
        ExcelComparisonOperator.GREATER_OR_EQUAL,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .LESS_THAN_OR_EQUAL)));

    assertEquals(
        ExcelDataValidationErrorStyle.STOP,
        ExcelDataValidationController.errorStyle(CTDataValidation.Factory.newInstance()));
    CTDataValidation warningStyle = CTDataValidation.Factory.newInstance();
    warningStyle.setErrorStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationErrorStyle.WARNING);
    assertEquals(
        ExcelDataValidationErrorStyle.WARNING,
        ExcelDataValidationController.errorStyle(warningStyle));
    CTDataValidation informationStyle = CTDataValidation.Factory.newInstance();
    informationStyle.setErrorStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationErrorStyle.INFORMATION);
    assertEquals(
        ExcelDataValidationErrorStyle.INFORMATION,
        ExcelDataValidationController.errorStyle(informationStyle));

    CTDataValidation invalidOperator = rawWholeOperator(null);
    try (var cursor = invalidOperator.newCursor()) {
      cursor.setAttributeText(new javax.xml.namespace.QName("", "operator"), "mystery");
    }
    IllegalArgumentException operatorFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelDataValidationController.comparisonOperator(invalidOperator));
    assertEquals("Unsupported raw validation operator: mystery", operatorFailure.getMessage());

    CTDataValidation invalidErrorStyle = CTDataValidation.Factory.newInstance();
    try (var cursor = invalidErrorStyle.newCursor()) {
      cursor.setAttributeText(new javax.xml.namespace.QName("", "errorStyle"), "mystery");
    }
    IllegalArgumentException errorStyleFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelDataValidationController.errorStyle(invalidErrorStyle));
    assertEquals("Unsupported raw validation error style: mystery", errorStyleFailure.getMessage());

    assertEquals(
        "between operator",
        ExcelDataValidationController.operatorLabel(ExcelComparisonOperator.BETWEEN));
    assertEquals(
        "not-between operator",
        ExcelDataValidationController.operatorLabel(ExcelComparisonOperator.NOT_BETWEEN));
    assertEquals(
        "equal operator",
        ExcelDataValidationController.operatorLabel(ExcelComparisonOperator.EQUAL));
    assertEquals(
        "not-equal operator",
        ExcelDataValidationController.operatorLabel(ExcelComparisonOperator.NOT_EQUAL));
    assertEquals(
        "greater-than operator",
        ExcelDataValidationController.operatorLabel(ExcelComparisonOperator.GREATER_THAN));
    assertEquals(
        "greater-or-equal operator",
        ExcelDataValidationController.operatorLabel(ExcelComparisonOperator.GREATER_OR_EQUAL));
    assertEquals(
        "less-than operator",
        ExcelDataValidationController.operatorLabel(ExcelComparisonOperator.LESS_THAN));
    assertEquals(
        "less-or-equal operator",
        ExcelDataValidationController.operatorLabel(ExcelComparisonOperator.LESS_OR_EQUAL));
  }

  @Test
  void directSnapshotDispatchCoversRemainingXssfAndRawComparisonFamilies() {
    ExcelDataValidationSnapshot.Supported decimalSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.DECIMAL,
                    ComparisonOperator.GT,
                    null,
                    "0.5",
                    null),
                List.of("A1")));
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", null),
        decimalSnapshot.validation().rule());

    ExcelDataValidationSnapshot.Supported dateSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.DATE,
                    ComparisonOperator.EQUAL,
                    null,
                    "DATE(2026,4,1)",
                    null),
                List.of("B1")));
    assertEquals(
        new ExcelDataValidationRule.DateRule(ExcelComparisonOperator.EQUAL, "DATE(2026,4,1)", null),
        dateSnapshot.validation().rule());

    ExcelDataValidationSnapshot.Supported timeSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.TIME,
                    ComparisonOperator.GT,
                    null,
                    "TIME(9,0,0)",
                    null),
                List.of("C1")));
    assertEquals(
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", null),
        timeSnapshot.validation().rule());

    ExcelDataValidationSnapshot.Supported textLengthSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.TEXT_LENGTH,
                    ComparisonOperator.LE,
                    null,
                    "20",
                    null),
                List.of("D1")));
    assertEquals(
        new ExcelDataValidationRule.TextLength(ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
        textLengthSnapshot.validation().rule());

    ExcelDataValidationSnapshot.Supported wholeBetweenSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.INTEGER,
                    ComparisonOperator.BETWEEN,
                    null,
                    "1",
                    "9"),
                List.of("E1")));
    assertEquals(
        new ExcelDataValidationRule.WholeNumber(ExcelComparisonOperator.BETWEEN, "1", "9"),
        wholeBetweenSnapshot.validation().rule());
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("F1"),
            "MISSING_FORMULA",
            "whole-number validation is missing formula2 for between operator."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER,
                ComparisonOperator.BETWEEN,
                null,
                "1",
                null),
            List.of("F1")));
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("G1"), "MISSING_FORMULA", "Custom formula validation is missing formula1."),
        ExcelDataValidationController.toSnapshot(
            rawValidation(STDataValidationType.CUSTOM, "G1", null), List.of("G1")));

    CTDataValidation allowBlankComparison = rawValidation(STDataValidationType.WHOLE, "H1", "1");
    allowBlankComparison.setOperator(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.NOT_BETWEEN);
    allowBlankComparison.setFormula2("9");
    allowBlankComparison.setAllowBlank(true);
    ExcelDataValidationSnapshot.Supported allowBlankComparisonSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(allowBlankComparison, List.of("H1")));
    assertTrue(allowBlankComparisonSnapshot.validation().allowBlank());

    CTDataValidation blankFormula2 = rawValidation(STDataValidationType.WHOLE, "I1", "1");
    blankFormula2.setOperator(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.BETWEEN);
    blankFormula2.setFormula2(" ");
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("I1"),
            "MISSING_FORMULA",
            "whole-number validation is missing formula2 for between operator."),
        ExcelDataValidationController.toSnapshot(blankFormula2, List.of("I1")));

    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("J1"), "MISSING_FORMULA", "whole-number validation is missing formula1."),
        ExcelDataValidationController.toSnapshot(
            rawValidation(STDataValidationType.WHOLE, "J1", null), List.of("J1")));

    CTDataValidation customAllowBlank =
        rawValidation(STDataValidationType.CUSTOM, "K1", "LEN(K1)>0");
    customAllowBlank.setAllowBlank(true);
    ExcelDataValidationSnapshot.Supported customAllowBlankSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(customAllowBlank, List.of("K1")));
    assertTrue(customAllowBlankSnapshot.validation().allowBlank());

    assertEquals(
        new ExcelDataValidationSnapshot.Supported(
            List.of("L1"),
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.FormulaList("\"Queued"), false, true, null, null)),
        ExcelDataValidationController.toSnapshot(
            rawValidation(STDataValidationType.LIST, "L1", "\"Queued"), List.of("L1")));

    CTDataValidation unknownType = CTDataValidation.Factory.newInstance();
    unknownType.setSqref(List.of("M1"));
    try (var cursor = unknownType.newCursor()) {
      cursor.setAttributeText(new javax.xml.namespace.QName("", "type"), "mystery");
    }
    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("M1"), "UNKNOWN", "Unsupported data-validation type: mystery"),
        ExcelDataValidationController.toSnapshot(unknownType, List.of("M1")));
  }

  @Test
  void rawCtSnapshotsMapSupportedFamiliesOperatorsAndMetadata() throws Exception {
    CTDataValidation explicitList = CTDataValidation.Factory.newInstance();
    explicitList.setType(STDataValidationType.LIST);
    explicitList.setSqref(List.of("A1:A3"));
    explicitList.setFormula1("\"Queued,Done\"");
    explicitList.setAllowBlank(true);
    explicitList.setShowDropDown(false);
    explicitList.setPromptTitle("Status");
    explicitList.setPrompt("Choose one workflow state.");
    explicitList.setShowInputMessage(true);
    explicitList.setErrorTitle("Invalid status");
    explicitList.setError("Use one of the allowed values.");
    explicitList.setShowErrorMessage(true);
    explicitList.setErrorStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationErrorStyle.WARNING);

    ExcelDataValidationSnapshot.Supported explicitListSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(explicitList, List.of("A1:A3")));
    assertEquals(
        new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
        explicitListSnapshot.validation().rule());
    assertEquals(
        new ExcelDataValidationPrompt("Status", "Choose one workflow state.", true),
        explicitListSnapshot.validation().prompt());
    assertEquals(
        ExcelDataValidationErrorStyle.WARNING,
        explicitListSnapshot.validation().errorAlert().style());
    assertTrue(explicitListSnapshot.validation().suppressDropDownArrow());

    CTDataValidation formulaList = CTDataValidation.Factory.newInstance();
    formulaList.setType(STDataValidationType.LIST);
    formulaList.setSqref(List.of("B1:B3"));
    formulaList.setFormula1("Statuses");
    formulaList.setShowDropDown(true);
    formulaList.setErrorTitle("Heads up");
    formulaList.setError("Read the allowed values.");
    formulaList.setShowErrorMessage(true);
    formulaList.setErrorStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationErrorStyle.INFORMATION);

    ExcelDataValidationSnapshot.Supported formulaListSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(formulaList, List.of("B1:B3")));
    assertEquals(
        new ExcelDataValidationRule.FormulaList("Statuses"),
        formulaListSnapshot.validation().rule());
    assertFalse(formulaListSnapshot.validation().suppressDropDownArrow());
    assertEquals(
        ExcelDataValidationErrorStyle.INFORMATION,
        formulaListSnapshot.validation().errorAlert().style());

    assertRawComparisonSnapshot(
        STDataValidationType.WHOLE,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.BETWEEN,
        ExcelComparisonOperator.BETWEEN);
    assertRawComparisonSnapshot(
        STDataValidationType.DECIMAL,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.NOT_BETWEEN,
        ExcelComparisonOperator.NOT_BETWEEN);
    assertRawComparisonSnapshot(
        STDataValidationType.DATE,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.EQUAL,
        ExcelComparisonOperator.NOT_EQUAL);
    assertRawComparisonSnapshot(
        STDataValidationType.TIME,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.NOT_EQUAL,
        ExcelComparisonOperator.EQUAL);
    assertRawComparisonSnapshot(
        STDataValidationType.TEXT_LENGTH,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.GREATER_THAN,
        ExcelComparisonOperator.LESS_THAN);
    assertRawComparisonSnapshot(
        STDataValidationType.WHOLE,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.LESS_THAN,
        ExcelComparisonOperator.GREATER_THAN);
    assertRawComparisonSnapshot(
        STDataValidationType.DECIMAL,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
            .GREATER_THAN_OR_EQUAL,
        ExcelComparisonOperator.LESS_OR_EQUAL);
    assertRawComparisonSnapshot(
        STDataValidationType.DATE,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
            .LESS_THAN_OR_EQUAL,
        ExcelComparisonOperator.GREATER_OR_EQUAL);

    CTDataValidation custom = CTDataValidation.Factory.newInstance();
    custom.setType(STDataValidationType.CUSTOM);
    custom.setSqref(List.of("C1:C3"));
    custom.setFormula1("LEN(C1)>0");
    ExcelDataValidationSnapshot.Supported customSnapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(custom, List.of("C1:C3")));
    assertEquals(
        new ExcelDataValidationRule.CustomFormula("LEN(C1)>0"), customSnapshot.validation().rule());
  }

  @Test
  void rawCtSnapshotsSurfaceUnknownOperators() throws Exception {
    CTDataValidation invalidOperator = CTDataValidation.Factory.newInstance();
    invalidOperator.setType(STDataValidationType.WHOLE);
    invalidOperator.setSqref(List.of("D1"));
    invalidOperator.setFormula1("1");
    try (var cursor = invalidOperator.newCursor()) {
      cursor.setAttributeText(new javax.xml.namespace.QName("", "operator"), "mystery");
    }

    assertEquals(
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("D1"),
            "UNKNOWN_OPERATOR",
            "whole-number validation uses an unsupported comparison operator."),
        ExcelDataValidationController.toSnapshot(invalidOperator, List.of("D1")));
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

  @Test
  void countAndHealthAnalysisHandleMalformedRawRulesWithoutUsingPoiWrappers() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, "E2", STDataValidationType.LIST, "\"\"");
      addRawValidation(sheet, "F2", STDataValidationType.LIST, null);
      addRawValidation(sheet, "G2", STDataValidationType.TEXT_LENGTH, "20");

      assertEquals(3, controller.dataValidationCount(sheet));
      List<ExcelDataValidationSnapshot> snapshots =
          controller.dataValidations(sheet, new ExcelRangeSelection.All());
      assertEquals(3, snapshots.size());
      assertEquals(
          new ExcelDataValidationSnapshot.Unsupported(
              List.of("E2"), "EMPTY_EXPLICIT_LIST", "Explicit list has no values."),
          snapshots.get(0));
      assertEquals(
          new ExcelDataValidationSnapshot.Unsupported(
              List.of("F2"),
              "MISSING_FORMULA",
              "List validation is missing both explicit values and formula1."),
          snapshots.get(1));
      ExcelDataValidationSnapshot.Supported textLength =
          assertInstanceOf(ExcelDataValidationSnapshot.Supported.class, snapshots.get(2));
      assertEquals(List.of("G2"), textLength.ranges());
      assertInstanceOf(ExcelDataValidationRule.TextLength.class, textLength.validation().rule());
      assertEquals(
          List.of(
              WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_EMPTY_EXPLICIT_LIST,
              WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_MALFORMED_RULE),
          controller.dataValidationHealthFindings("Budget", sheet).stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .distinct()
              .toList());
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

  private void assertRawComparisonSnapshot(
      STDataValidationType.Enum type,
      org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.Enum operator,
      ExcelComparisonOperator expectedOperator) {
    CTDataValidation validation = CTDataValidation.Factory.newInstance();
    validation.setType(type);
    validation.setSqref(List.of("Z1"));
    validation.setFormula1("1");
    if (operator
            == org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.BETWEEN
        || operator
            == org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                .NOT_BETWEEN) {
      validation.setFormula2("2");
    }
    validation.setOperator(operator);

    ExcelDataValidationSnapshot.Supported snapshot =
        assertInstanceOf(
            ExcelDataValidationSnapshot.Supported.class,
            ExcelDataValidationController.toSnapshot(validation, List.of("Z1")));

    ExcelDataValidationRule rule = snapshot.validation().rule();
    ExcelComparisonOperator actualOperator =
        switch (rule) {
          case ExcelDataValidationRule.WholeNumber wholeNumber -> wholeNumber.operator();
          case ExcelDataValidationRule.DecimalNumber decimalNumber -> decimalNumber.operator();
          case ExcelDataValidationRule.DateRule dateRule -> dateRule.operator();
          case ExcelDataValidationRule.TimeRule timeRule -> timeRule.operator();
          case ExcelDataValidationRule.TextLength textLength -> textLength.operator();
          default -> throw new AssertionError("unexpected rule type: " + rule);
        };
    assertEquals(expectedOperator, actualOperator);
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
    addRawValidation(sheet, range, type, null);
  }

  private static void addRawValidation(
      XSSFSheet sheet, String range, STDataValidationType.Enum type, String formula1) {
    CTDataValidations dataValidations =
        sheet.getCTWorksheet().isSetDataValidations()
            ? sheet.getCTWorksheet().getDataValidations()
            : sheet.getCTWorksheet().addNewDataValidations();
    CTDataValidation validation = dataValidations.addNewDataValidation();
    validation.setType(type);
    validation.setSqref(List.of(range));
    if (formula1 != null) {
      validation.setFormula1(formula1);
    }
    if (type == STDataValidationType.TEXT_LENGTH) {
      validation.setOperator(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.LESS_THAN);
    }
    syncCount(sheet);
  }

  private static CTDataValidation rawValidation(
      STDataValidationType.Enum type, String range, String formula1) {
    CTDataValidation validation = CTDataValidation.Factory.newInstance();
    validation.setType(type);
    validation.setSqref(List.of(range));
    if (formula1 != null) {
      validation.setFormula1(formula1);
    }
    return validation;
  }

  private static CTDataValidation rawWholeOperator(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.Enum operator) {
    CTDataValidation validation = rawValidation(STDataValidationType.WHOLE, "Z1", "1");
    if (operator != null) {
      validation.setOperator(operator);
    }
    return validation;
  }

  private static StubXssfDataValidation xssfListStub() {
    return new StubXssfDataValidation(
        new StubConstraint(
            DataValidationConstraint.ValidationType.LIST, 0, new String[] {"Queued"}, null, null));
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
