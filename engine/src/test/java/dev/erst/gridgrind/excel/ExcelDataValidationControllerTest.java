package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationController;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
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
        new ResolvedDataValidation(
            new DataValidationRuleSpec.FormulaList(
                ExcelComparisonFormulaSupport.normalizeFormula("=Statuses", "formula")),
            false,
            true,
            Optional.empty(),
            Optional.empty()));
  }

  @Test
  void roundTripsWholeNumberValidation() throws IOException {
    assertRoundTrip(
        "C1:C3",
        new ResolvedDataValidation(
            new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.GREATER_THAN,
                ExcelComparisonFormulaSupport.normalizeFormula("=1", "formula1"),
                ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                    Optional.empty(), "formula2")),
            false,
            false,
            Optional.empty(),
            Optional.empty()));
  }

  @Test
  void roundTripsDecimalValidation() throws IOException {
    assertRoundTrip(
        "D1:D3",
        new ResolvedDataValidation(
            new DataValidationRuleSpec.DecimalNumber(
                ExcelComparisonOperator.GREATER_THAN,
                ExcelComparisonFormulaSupport.normalizeFormula("=0.5", "formula1"),
                ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                    Optional.empty(), "formula2")),
            false,
            false,
            Optional.empty(),
            Optional.of(warningAlert())));
  }

  @Test
  void roundTripsDateValidation() throws IOException {
    assertRoundTrip(
        "E1:E3",
        new ResolvedDataValidation(
            new DataValidationRuleSpec.DateRule(
                ExcelComparisonOperator.EQUAL,
                ExcelComparisonFormulaSupport.normalizeFormula("=DATE(2026,4,1)", "formula1"),
                ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                    Optional.empty(), "formula2")),
            false,
            false,
            Optional.empty(),
            Optional.empty()));
  }

  @Test
  void roundTripsTimeValidation() throws IOException {
    assertRoundTrip(
        "F1:F3",
        new ResolvedDataValidation(
            new DataValidationRuleSpec.TimeRule(
                ExcelComparisonOperator.GREATER_THAN,
                ExcelComparisonFormulaSupport.normalizeFormula("=TIME(9,0,0)", "formula1"),
                ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                    Optional.empty(), "formula2")),
            false,
            false,
            Optional.empty(),
            Optional.empty()));
  }

  @Test
  void roundTripsTextLengthValidation() throws IOException {
    assertRoundTrip(
        "G1:G3",
        new ResolvedDataValidation(
            new DataValidationRuleSpec.TextLength(
                ExcelComparisonOperator.GREATER_THAN,
                ExcelComparisonFormulaSupport.normalizeFormula("=20", "formula1"),
                ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                    Optional.empty(), "formula2")),
            false,
            false,
            Optional.empty(),
            Optional.empty()));
  }

  @Test
  void roundTripsCustomFormulaValidation() throws IOException {
    assertRoundTrip(
        "H1:H3",
        new ResolvedDataValidation(
            new DataValidationRuleSpec.CustomFormula(
                ExcelComparisonFormulaSupport.normalizeFormula("=LEN(H1)>0", "formula")),
            false,
            false,
            Optional.empty(),
            Optional.empty()));
  }

  @Test
  void selectedSelectionFiltersByIntersection() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, "A1:A3", explicitListValidation());
      controller.setDataValidation(
          sheet,
          "C1:C3",
          new ResolvedDataValidation(
              new DataValidationRuleSpec.TextLength(
                  ExcelComparisonOperator.GREATER_THAN,
                  ExcelComparisonFormulaSupport.normalizeFormula("20", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      Optional.empty(), "formula2")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));

      List<DataValidationEntryReport> selected =
          controller.dataValidations(
              sheet,
              List.of("C2").stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
                  .toList());

      assertEquals(
          List.of(
              new DataValidationEntryReport.Supported(
                  List.of("C1:C3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  dev.erst.gridgrind.excel.DataValidationFixtures.report(
                      expectedRoundTripDefinition(
                          new ResolvedDataValidation(
                              new DataValidationRuleSpec.TextLength(
                                  ExcelComparisonOperator.GREATER_THAN,
                                  ExcelComparisonFormulaSupport.normalizeFormula("20", "formula1"),
                                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                                      Optional.empty(), "formula2")),
                              false,
                              false,
                              Optional.empty(),
                              Optional.empty()))))),
          selected);
    }
  }

  @Test
  void clearDataValidationsAllRemovesEveryStructure() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, "A1:A3", explicitListValidation());
      controller.setDataValidation(sheet, "B1:B3", explicitListValidation());

      controller.clearDataValidations(sheet);

      assertEquals(0, controller.dataValidationCount(sheet));
      assertFalse(sheet.getCTWorksheet().isSetDataValidations());
    }
  }

  @Test
  void clearDataValidationsAllOnEmptySheetDoesNothing() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.clearDataValidations(sheet);

      assertEquals(0, controller.dataValidationCount(sheet));
      assertFalse(sheet.getCTWorksheet().isSetDataValidations());
    }
  }

  @Test
  void selectedClearOnEmptySheetDoesNothing() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.clearDataValidations(
          sheet,
          List.of("A1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList());

      assertEquals(0, controller.dataValidationCount(sheet));
      assertFalse(sheet.getCTWorksheet().isSetDataValidations());
    }
  }

  @Test
  void selectedClearRemovesWholeStructuresAndUnsetsTheContainer() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, "A1:A3", explicitListValidation());

      controller.clearDataValidations(
          sheet,
          List.of("A1:A3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList());

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

      controller.clearDataValidations(
          sheet,
          List.of("A3:A4").stream().map(dev.erst.gridgrind.core.vocabulary.RangeRef::new).toList());

      assertEquals(
          List.of(
              new DataValidationEntryReport.Supported(
                  List.of("A1:A2", "A5:A10").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  dev.erst.gridgrind.excel.DataValidationFixtures.report(
                      expectedRoundTripDefinition(explicitListValidation())))),
          controller.dataValidations(sheet));
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
          List.of("B20:B20", "A30:A31", "A1:B3", "D20:D20").stream()
              .map(dev.erst.gridgrind.core.vocabulary.RangeRef::new)
              .toList());

      assertEquals(
          List.of(
              new DataValidationEntryReport.Supported(
                  List.of("A20", "C20").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  dev.erst.gridgrind.excel.DataValidationFixtures.report(
                      expectedRoundTripDefinition(explicitListValidation()))),
              new DataValidationEntryReport.Supported(
                  List.of("C1:C3").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  dev.erst.gridgrind.excel.DataValidationFixtures.report(
                      expectedRoundTripDefinition(explicitListValidation())))),
          controller.dataValidations(sheet));
    }
  }

  @Test
  void dataValidationsExposeUnsupportedStructures() {
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("A1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "ANY",
            "Excel 'any value' validation is not modeled by GridGrind."),
        invokeToSnapshot(
            new StubConstraint(DataValidationConstraint.ValidationType.ANY, 0, null, null, null),
            List.of("A1")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("B1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "UNKNOWN_OPERATOR",
            "whole-number validation uses an unsupported comparison operator."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER, -99, null, "1", null),
            List.of("B1")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("B2").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "UNKNOWN",
            "Unsupported data-validation type: 999"),
        invokeToSnapshot(new StubConstraint(999, 0, null, "1", null), List.of("B2")));
    assertEquals(
        new DataValidationEntryReport.Supported(
            List.of("C0").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            dev.erst.gridgrind.excel.DataValidationFixtures.report(
                new ResolvedDataValidation(
                    new DataValidationRuleSpec.ExplicitList(List.of()),
                    false,
                    true,
                    Optional.empty(),
                    Optional.empty()))),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.LIST, 0, new String[0], null, null),
            List.of("C0")));
    assertEquals(
        new DataValidationEntryReport.Supported(
            List.of("C0b").stream()
                .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                .toList(),
            dev.erst.gridgrind.excel.DataValidationFixtures.report(
                new ResolvedDataValidation(
                    new DataValidationRuleSpec.ExplicitList(List.of()),
                    false,
                    true,
                    Optional.empty(),
                    Optional.empty()))),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.LIST, 0, new String[0], "\"\"", null),
            List.of("C0b")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("C1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "List validation is missing both explicit values and formula1."),
        invokeToSnapshot(
            new StubConstraint(DataValidationConstraint.ValidationType.LIST, 0, null, null, null),
            List.of("C1")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("C2").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "List validation is missing both explicit values and formula1."),
        invokeToSnapshot(
            new StubConstraint(DataValidationConstraint.ValidationType.LIST, 0, null, " ", null),
            List.of("C2")));
    assertEquals(
        new DataValidationEntryReport.Supported(
            List.of("C3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            dev.erst.gridgrind.excel.DataValidationFixtures.report(
                new ResolvedDataValidation(
                    new DataValidationRuleSpec.FormulaList(
                        ExcelComparisonFormulaSupport.normalizeFormula("Statuses", "formula")),
                    false,
                    true,
                    Optional.empty(),
                    Optional.empty()))),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.LIST, 0, null, "Statuses", null),
            List.of("C3")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("D1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "whole-number validation is missing formula1."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.GREATER_THAN,
                null,
                null,
                null),
            List.of("D1")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("D2").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "whole-number validation is missing formula1."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.GREATER_THAN,
                null,
                " ",
                null),
            List.of("D2")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("E1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "Custom formula validation is missing formula1."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.FORMULA, 0, null, null, null),
            List.of("E1")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("E2").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "Custom formula validation is missing formula1."),
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

    DataValidationEntryReport.Supported blankPromptTitleSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            ExcelDataValidationController.toSnapshot(blankPromptTitle, List.of("A1")));
    assertEquals(Optional.empty(), blankPromptTitleSnapshot.validation().prompt());
    assertEquals(Optional.empty(), blankPromptTitleSnapshot.validation().errorAlert());

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

    DataValidationEntryReport.Supported blankPromptTextSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            ExcelDataValidationController.toSnapshot(blankPromptText, List.of("B1")));
    assertEquals(Optional.empty(), blankPromptTextSnapshot.validation().prompt());
    assertEquals(Optional.empty(), blankPromptTextSnapshot.validation().errorAlert());
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

    DataValidationEntryReport.Supported snapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            ExcelDataValidationController.toSnapshot(nullPromptText, List.of("C1")));

    assertEquals(Optional.empty(), snapshot.validation().prompt());
    assertEquals(Optional.empty(), snapshot.validation().errorAlert());
  }

  @Test
  void xssfMetadataHelpersCoverPromptErrorAndSupportedFormulaPayloads() {
    assertEquals(
        Optional.empty(),
        ExcelDataValidationController.prompt(
            xssfListStub().overridePrompt(null, "Choose one workflow state.", true)));
    assertEquals(
        Optional.empty(),
        ExcelDataValidationController.prompt(
            xssfListStub().overridePrompt(" ", "Choose one workflow state.", true)));
    assertEquals(
        Optional.empty(),
        ExcelDataValidationController.prompt(xssfListStub().overridePrompt("Status", null, true)));
    assertEquals(
        Optional.empty(),
        ExcelDataValidationController.prompt(xssfListStub().overridePrompt("Status", " ", true)));
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationPromptReport(
                "Status", "Choose one workflow state.", false)),
        ExcelDataValidationController.prompt(
            xssfListStub().overridePrompt("Status", "Choose one workflow state.", false)));

    assertEquals(
        Optional.empty(),
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError(
                    null, "Use an allowed value.", true, DataValidation.ErrorStyle.STOP)));
    assertEquals(
        Optional.empty(),
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError(
                    " ", "Use an allowed value.", true, DataValidation.ErrorStyle.STOP)));
    assertEquals(
        Optional.empty(),
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError("Invalid status", null, true, DataValidation.ErrorStyle.STOP)));
    assertEquals(
        Optional.empty(),
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError("Invalid status", " ", true, DataValidation.ErrorStyle.STOP)));
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport(
                ExcelDataValidationErrorStyle.INFORMATION,
                "Heads up",
                "Use an allowed value.",
                false)),
        ExcelDataValidationController.errorAlert(
            xssfListStub()
                .overrideError(
                    "Heads up", "Use an allowed value.", false, DataValidation.ErrorStyle.INFO)));

    DataValidationEntryReport.Supported formulaListSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
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
        new DataValidationRuleSpec.FormulaList(
            ExcelComparisonFormulaSupport.normalizeFormula("Statuses", "formula")),
        formulaListSnapshot.validation().rule());
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationPromptReport(
                "Status", "Choose one workflow state.", false)),
        formulaListSnapshot.validation().prompt());
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport(
                ExcelDataValidationErrorStyle.WARNING,
                "Heads up",
                "Read the allowed values.",
                false)),
        formulaListSnapshot.validation().errorAlert());

    DataValidationEntryReport.Supported customFormulaSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
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
        new DataValidationRuleSpec.CustomFormula(
            ExcelComparisonFormulaSupport.normalizeFormula("LEN(D1)>0", "formula")),
        customFormulaSnapshot.validation().rule());
  }

  @Test
  void rawCtMetadataHelpersCoverPromptAndErrorBranches() {
    CTDataValidation promptDefaults = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    promptDefaults.setPromptTitle("Status");
    promptDefaults.setPrompt("Choose one workflow state.");
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationPromptReport(
                "Status", "Choose one workflow state.", false)),
        ExcelDataValidationController.prompt(promptDefaults));

    CTDataValidation promptExplicitFalse =
        rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    promptExplicitFalse.setPromptTitle("Status");
    promptExplicitFalse.setPrompt("Choose one workflow state.");
    promptExplicitFalse.setShowInputMessage(false);
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationPromptReport(
                "Status", "Choose one workflow state.", false)),
        ExcelDataValidationController.prompt(promptExplicitFalse));

    CTDataValidation promptBlankTitle =
        rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    promptBlankTitle.setPromptTitle(" ");
    promptBlankTitle.setPrompt("Choose one workflow state.");
    assertEquals(Optional.empty(), ExcelDataValidationController.prompt(promptBlankTitle));

    CTDataValidation promptMissingText =
        rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    promptMissingText.setPromptTitle("Status");
    assertEquals(Optional.empty(), ExcelDataValidationController.prompt(promptMissingText));

    CTDataValidation promptBlankText = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    promptBlankText.setPromptTitle("Status");
    promptBlankText.setPrompt(" ");
    assertEquals(Optional.empty(), ExcelDataValidationController.prompt(promptBlankText));

    CTDataValidation errorDefaults = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    errorDefaults.setErrorTitle("Invalid status");
    errorDefaults.setError("Use one of the allowed values.");
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport(
                ExcelDataValidationErrorStyle.STOP,
                "Invalid status",
                "Use one of the allowed values.",
                false)),
        ExcelDataValidationController.errorAlert(errorDefaults));

    CTDataValidation warningError = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    warningError.setErrorTitle("Warning");
    warningError.setError("Watch this value.");
    warningError.setShowErrorMessage(true);
    warningError.setErrorStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationErrorStyle.WARNING);
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport(
                ExcelDataValidationErrorStyle.WARNING, "Warning", "Watch this value.", true)),
        ExcelDataValidationController.errorAlert(warningError));

    CTDataValidation infoError = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    infoError.setErrorTitle("Info");
    infoError.setError("Read the list.");
    infoError.setErrorStyle(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationErrorStyle.INFORMATION);
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport(
                ExcelDataValidationErrorStyle.INFORMATION, "Info", "Read the list.", false)),
        ExcelDataValidationController.errorAlert(infoError));

    CTDataValidation explicitFalseError =
        rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    explicitFalseError.setErrorTitle("Explicit false");
    explicitFalseError.setError("Stay false.");
    explicitFalseError.setShowErrorMessage(false);
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport(
                ExcelDataValidationErrorStyle.STOP, "Explicit false", "Stay false.", false)),
        ExcelDataValidationController.errorAlert(explicitFalseError));

    CTDataValidation blankErrorTitle = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    blankErrorTitle.setErrorTitle(" ");
    blankErrorTitle.setError("Use one of the allowed values.");
    assertEquals(Optional.empty(), ExcelDataValidationController.errorAlert(blankErrorTitle));

    CTDataValidation missingErrorText =
        rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    missingErrorText.setErrorTitle("Invalid status");
    assertEquals(Optional.empty(), ExcelDataValidationController.errorAlert(missingErrorText));

    CTDataValidation blankErrorText = rawValidation(STDataValidationType.LIST, "A1", "\"Queued\"");
    blankErrorText.setErrorTitle("Invalid status");
    blankErrorText.setError(" ");
    assertEquals(Optional.empty(), ExcelDataValidationController.errorAlert(blankErrorText));
  }

  @Test
  void rawEnumMappingsAndLabelsCoverDefaultedAndMalformedMetadata() {
    CTDataValidation defaultWhole = rawValidation(STDataValidationType.WHOLE, "B1", "1");
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("B1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
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
        ExcelComparisonOperator.EQUAL,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .EQUAL)));
    assertEquals(
        ExcelComparisonOperator.NOT_EQUAL,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .NOT_EQUAL)));
    assertEquals(
        ExcelComparisonOperator.GREATER_THAN,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .GREATER_THAN)));
    assertEquals(
        ExcelComparisonOperator.LESS_THAN,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .LESS_THAN)));
    assertEquals(
        ExcelComparisonOperator.GREATER_OR_EQUAL,
        ExcelDataValidationController.comparisonOperator(
            rawWholeOperator(
                org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
                    .GREATER_THAN_OR_EQUAL)));
    assertEquals(
        ExcelComparisonOperator.LESS_OR_EQUAL,
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
    assertEquals(
        "Unsupported OOXML data-validation comparison operator: mystery",
        operatorFailure.getMessage());

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
    DataValidationEntryReport.Supported decimalSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.DECIMAL,
                    DataValidationConstraint.OperatorType.GREATER_THAN,
                    null,
                    "0.5",
                    null),
                List.of("A1")));
    assertEquals(
        new DataValidationRuleSpec.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN,
            ExcelComparisonFormulaSupport.normalizeFormula("0.5", "formula1"),
            ExcelComparisonFormulaSupport.normalizeOptionalFormula(Optional.empty(), "formula2")),
        decimalSnapshot.validation().rule());

    DataValidationEntryReport.Supported dateSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.DATE,
                    DataValidationConstraint.OperatorType.EQUAL,
                    null,
                    "DATE(2026,4,1)",
                    null),
                List.of("B1")));
    assertEquals(
        new DataValidationRuleSpec.DateRule(
            ExcelComparisonOperator.EQUAL,
            ExcelComparisonFormulaSupport.normalizeFormula("DATE(2026,4,1)", "formula1"),
            ExcelComparisonFormulaSupport.normalizeOptionalFormula(Optional.empty(), "formula2")),
        dateSnapshot.validation().rule());

    DataValidationEntryReport.Supported timeSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.TIME,
                    DataValidationConstraint.OperatorType.GREATER_THAN,
                    null,
                    "TIME(9,0,0)",
                    null),
                List.of("C1")));
    assertEquals(
        new DataValidationRuleSpec.TimeRule(
            ExcelComparisonOperator.GREATER_THAN,
            ExcelComparisonFormulaSupport.normalizeFormula("TIME(9,0,0)", "formula1"),
            ExcelComparisonFormulaSupport.normalizeOptionalFormula(Optional.empty(), "formula2")),
        timeSnapshot.validation().rule());

    DataValidationEntryReport.Supported textLengthSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.TEXT_LENGTH,
                    DataValidationConstraint.OperatorType.LESS_OR_EQUAL,
                    null,
                    "20",
                    null),
                List.of("D1")));
    assertEquals(
        new DataValidationRuleSpec.TextLength(
            ExcelComparisonOperator.LESS_OR_EQUAL,
            ExcelComparisonFormulaSupport.normalizeFormula("20", "formula1"),
            ExcelComparisonFormulaSupport.normalizeOptionalFormula(Optional.empty(), "formula2")),
        textLengthSnapshot.validation().rule());

    DataValidationEntryReport.Supported wholeBetweenSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            invokeToSnapshot(
                new StubConstraint(
                    DataValidationConstraint.ValidationType.INTEGER,
                    DataValidationConstraint.OperatorType.BETWEEN,
                    null,
                    "1",
                    "9"),
                List.of("E1")));
    assertEquals(
        new DataValidationRuleSpec.WholeNumber(
            ExcelComparisonOperator.BETWEEN,
            ExcelComparisonFormulaSupport.normalizeFormula("1", "formula1"),
            ExcelComparisonFormulaSupport.normalizeOptionalFormula(Optional.of("9"), "formula2")),
        wholeBetweenSnapshot.validation().rule());
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("F1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "whole-number validation is missing formula2 for between operator."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.BETWEEN,
                null,
                "1",
                null),
            List.of("F1")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("F2").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "whole-number validation is missing formula2 for between operator."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.BETWEEN,
                null,
                "1",
                " "),
            List.of("F2")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("G1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "Custom formula validation is missing formula1."),
        ExcelDataValidationController.toSnapshot(
            rawValidation(STDataValidationType.CUSTOM, "G1", null), List.of("G1")));
    // Blank formula1 (non-null, whitespace-only) for XSSF list and comparison paths.
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("G2").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "List validation is missing both explicit values and formula1."),
        invokeToSnapshot(
            new StubConstraint(DataValidationConstraint.ValidationType.LIST, 0, null, " ", null),
            List.of("G2")));
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("G3").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "whole-number validation is missing formula1."),
        invokeToSnapshot(
            new StubConstraint(
                DataValidationConstraint.ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.GREATER_THAN,
                null,
                " ",
                null),
            List.of("G3")));

    CTDataValidation allowBlankComparison = rawValidation(STDataValidationType.WHOLE, "H1", "1");
    allowBlankComparison.setOperator(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.NOT_BETWEEN);
    allowBlankComparison.setFormula2("9");
    allowBlankComparison.setAllowBlank(true);
    DataValidationEntryReport.Supported allowBlankComparisonSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            ExcelDataValidationController.toSnapshot(allowBlankComparison, List.of("H1")));
    assertTrue(allowBlankComparisonSnapshot.validation().allowBlank());

    CTDataValidation blankFormula2 = rawValidation(STDataValidationType.WHOLE, "I1", "1");
    blankFormula2.setOperator(
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.BETWEEN);
    blankFormula2.setFormula2(" ");
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("I1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "whole-number validation is missing formula2 for between operator."),
        ExcelDataValidationController.toSnapshot(blankFormula2, List.of("I1")));

    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("J1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "MISSING_FORMULA",
            "whole-number validation is missing formula1."),
        ExcelDataValidationController.toSnapshot(
            rawValidation(STDataValidationType.WHOLE, "J1", null), List.of("J1")));

    CTDataValidation customAllowBlank =
        rawValidation(STDataValidationType.CUSTOM, "K1", "LEN(K1)>0");
    customAllowBlank.setAllowBlank(true);
    DataValidationEntryReport.Supported customAllowBlankSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            ExcelDataValidationController.toSnapshot(customAllowBlank, List.of("K1")));
    assertTrue(customAllowBlankSnapshot.validation().allowBlank());

    assertEquals(
        new DataValidationEntryReport.Supported(
            List.of("L1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            dev.erst.gridgrind.excel.DataValidationFixtures.report(
                new ResolvedDataValidation(
                    new DataValidationRuleSpec.FormulaList(
                        ExcelComparisonFormulaSupport.normalizeFormula("\"Queued", "formula")),
                    false,
                    true,
                    Optional.empty(),
                    Optional.empty()))),
        ExcelDataValidationController.toSnapshot(
            rawValidation(STDataValidationType.LIST, "L1", "\"Queued"), List.of("L1")));
    assertEquals(
        new DataValidationEntryReport.Supported(
            List.of("L2").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            dev.erst.gridgrind.excel.DataValidationFixtures.report(
                new ResolvedDataValidation(
                    new DataValidationRuleSpec.FormulaList(
                        ExcelComparisonFormulaSupport.normalizeFormula("\"", "formula")),
                    false,
                    true,
                    Optional.empty(),
                    Optional.empty()))),
        ExcelDataValidationController.toSnapshot(
            rawValidation(STDataValidationType.LIST, "L2", "\""), List.of("L2")));

    CTDataValidation unknownType = CTDataValidation.Factory.newInstance();
    unknownType.setSqref(List.of("M1"));
    try (var cursor = unknownType.newCursor()) {
      cursor.setAttributeText(new javax.xml.namespace.QName("", "type"), "mystery");
    }
    assertEquals(
        new DataValidationEntryReport.Unsupported(
            List.of("M1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
            "UNKNOWN",
            "Unsupported data-validation type: mystery"),
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

    DataValidationEntryReport.Supported explicitListSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            ExcelDataValidationController.toSnapshot(explicitList, List.of("A1:A3")));
    assertEquals(
        new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
        explicitListSnapshot.validation().rule());
    assertEquals(
        Optional.of(
            new dev.erst.gridgrind.core.model.validation.DataValidationPromptReport(
                "Status", "Choose one workflow state.", true)),
        explicitListSnapshot.validation().prompt());
    assertEquals(
        ExcelDataValidationErrorStyle.WARNING,
        explicitListSnapshot.validation().errorAlert().orElseThrow().style());
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

    DataValidationEntryReport.Supported formulaListSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            ExcelDataValidationController.toSnapshot(formulaList, List.of("B1:B3")));
    assertEquals(
        new DataValidationRuleSpec.FormulaList(
            ExcelComparisonFormulaSupport.normalizeFormula("Statuses", "formula")),
        formulaListSnapshot.validation().rule());
    assertFalse(formulaListSnapshot.validation().suppressDropDownArrow());
    assertEquals(
        ExcelDataValidationErrorStyle.INFORMATION,
        formulaListSnapshot.validation().errorAlert().orElseThrow().style());

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
        ExcelComparisonOperator.EQUAL);
    assertRawComparisonSnapshot(
        STDataValidationType.TIME,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.NOT_EQUAL,
        ExcelComparisonOperator.NOT_EQUAL);
    assertRawComparisonSnapshot(
        STDataValidationType.TEXT_LENGTH,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.GREATER_THAN,
        ExcelComparisonOperator.GREATER_THAN);
    assertRawComparisonSnapshot(
        STDataValidationType.WHOLE,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator.LESS_THAN,
        ExcelComparisonOperator.LESS_THAN);
    assertRawComparisonSnapshot(
        STDataValidationType.DECIMAL,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
            .GREATER_THAN_OR_EQUAL,
        ExcelComparisonOperator.GREATER_OR_EQUAL);
    assertRawComparisonSnapshot(
        STDataValidationType.DATE,
        org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationOperator
            .LESS_THAN_OR_EQUAL,
        ExcelComparisonOperator.LESS_OR_EQUAL);

    CTDataValidation custom = CTDataValidation.Factory.newInstance();
    custom.setType(STDataValidationType.CUSTOM);
    custom.setSqref(List.of("C1:C3"));
    custom.setFormula1("LEN(C1)>0");
    DataValidationEntryReport.Supported customSnapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            ExcelDataValidationController.toSnapshot(custom, List.of("C1:C3")));
    assertEquals(
        new DataValidationRuleSpec.CustomFormula(
            ExcelComparisonFormulaSupport.normalizeFormula("LEN(C1)>0", "formula")),
        customSnapshot.validation().rule());
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
        new DataValidationEntryReport.Unsupported(
            List.of("D1").stream().map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new).toList(),
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
          new ResolvedDataValidation(
              new DataValidationRuleSpec.FormulaList(
                  ExcelComparisonFormulaSupport.normalizeFormula("#REF!", "formula")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));
      controller.setDataValidation(
          sheet,
          "B1:B3",
          new ResolvedDataValidation(
              new DataValidationRuleSpec.WholeNumber(
                  ExcelComparisonOperator.BETWEEN,
                  ExcelComparisonFormulaSupport.normalizeFormula("#REF!", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      Optional.of("#REF!"), "formula2")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));
      controller.setDataValidation(
          sheet,
          "C1:C3",
          new ResolvedDataValidation(
              new DataValidationRuleSpec.DecimalNumber(
                  ExcelComparisonOperator.GREATER_THAN,
                  ExcelComparisonFormulaSupport.normalizeFormula("#REF!", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      Optional.empty(), "formula2")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));
      controller.setDataValidation(
          sheet,
          "D1:D3",
          new ResolvedDataValidation(
              new DataValidationRuleSpec.DateRule(
                  ExcelComparisonOperator.GREATER_THAN,
                  ExcelComparisonFormulaSupport.normalizeFormula("#REF!", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      Optional.empty(), "formula2")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));
      controller.setDataValidation(
          sheet,
          "E1:E3",
          new ResolvedDataValidation(
              new DataValidationRuleSpec.TimeRule(
                  ExcelComparisonOperator.GREATER_THAN,
                  ExcelComparisonFormulaSupport.normalizeFormula("#REF!", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      Optional.empty(), "formula2")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));
      controller.setDataValidation(
          sheet,
          "F1:F3",
          new ResolvedDataValidation(
              new DataValidationRuleSpec.TextLength(
                  ExcelComparisonOperator.GREATER_THAN,
                  ExcelComparisonFormulaSupport.normalizeFormula("#REF!", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      Optional.empty(), "formula2")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));
      controller.setDataValidation(
          sheet,
          "G1:G3",
          new ResolvedDataValidation(
              new DataValidationRuleSpec.CustomFormula(
                  ExcelComparisonFormulaSupport.normalizeFormula("#REF!", "formula")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));
      addHelperValidation(sheet, "H1:H3", helperValidation(sheet, 0, 2, 7, 7));
      addHelperValidation(sheet, "H2:H4", helperValidation(sheet, 1, 3, 7, 7));

      List<AnalysisFindingReport> findings =
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
          new ResolvedDataValidation(
              new DataValidationRuleSpec.WholeNumber(
                  ExcelComparisonOperator.GREATER_THAN,
                  ExcelComparisonFormulaSupport.normalizeFormula("1", "formula1"),
                  ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      Optional.empty(), "formula2")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));
      controller.setDataValidation(
          sheet,
          "B1:B3",
          new ResolvedDataValidation(
              new DataValidationRuleSpec.FormulaList(
                  ExcelComparisonFormulaSupport.normalizeFormula("Statuses", "formula")),
              false,
              true,
              Optional.empty(),
              Optional.empty()));
      controller.setDataValidation(
          sheet,
          "C1:C3",
          new ResolvedDataValidation(
              new DataValidationRuleSpec.CustomFormula(
                  ExcelComparisonFormulaSupport.normalizeFormula("LEN(C1)>0", "formula")),
              false,
              false,
              Optional.empty(),
              Optional.empty()));
      addRawValidation(sheet, "D1", STDataValidationType.NONE);

      List<AnalysisFindingReport> findings =
          controller.dataValidationHealthFindings("Budget", sheet);

      assertEquals(4, controller.dataValidationCount(sheet));
      assertEquals(
          List.of(AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE),
          findings.stream().map(AnalysisFindingReport::code).toList());
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
      List<DataValidationEntryReport> snapshots = controller.dataValidations(sheet);
      assertEquals(3, snapshots.size());
      DataValidationEntryReport.Supported emptyExplicitList =
          assertInstanceOf(DataValidationEntryReport.Supported.class, snapshots.get(0));
      assertEquals(
          List.of("E2"),
          emptyExplicitList.ranges().stream()
              .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
              .toList());
      assertEquals(
          new DataValidationRuleSpec.ExplicitList(List.of()),
          emptyExplicitList.validation().rule());
      assertEquals(
          new DataValidationEntryReport.Unsupported(
              List.of("F2").stream()
                  .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                  .toList(),
              "MISSING_FORMULA",
              "List validation is missing both explicit values and formula1."),
          snapshots.get(1));
      DataValidationEntryReport.Supported textLength =
          assertInstanceOf(DataValidationEntryReport.Supported.class, snapshots.get(2));
      assertEquals(
          List.of("G2"),
          textLength.ranges().stream()
              .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
              .toList());
      assertInstanceOf(DataValidationRuleSpec.TextLength.class, textLength.validation().rule());
      assertEquals(
          List.of(
              AnalysisFindingCode.DATA_VALIDATION_EMPTY_EXPLICIT_LIST,
              AnalysisFindingCode.DATA_VALIDATION_MALFORMED_RULE),
          controller.dataValidationHealthFindings("Budget", sheet).stream()
              .map(AnalysisFindingReport::code)
              .distinct()
              .toList());
    }
  }

  @Test
  void healthAnalysisMapsMissingFormulaAndUnknownKindsDistinctly() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      addRawValidation(sheet, "A1", STDataValidationType.LIST, null);
      addRawValidation(sheet, "B1", STDataValidationType.NONE);

      List<AnalysisFindingCode> codes =
          controller.dataValidationHealthFindings("Budget", sheet).stream()
              .map(AnalysisFindingReport::code)
              .distinct()
              .toList();

      assertEquals(
          List.of(
              AnalysisFindingCode.DATA_VALIDATION_MALFORMED_RULE,
              AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE),
          codes);
    }
  }

  @Test
  void setDataValidationRoundTripsEmptyExplicitListsAndReportsHealth() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      ResolvedDataValidation definition =
          new ResolvedDataValidation(
              new DataValidationRuleSpec.ExplicitList(List.of()),
              true,
              false,
              Optional.empty(),
              Optional.empty());

      controller.setDataValidation(sheet, "E2:E4", definition);

      List<DataValidationEntryReport> snapshots = controller.dataValidations(sheet);

      assertEquals(
          List.of(
              new DataValidationEntryReport.Supported(
                  List.of("E2:E4").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  dev.erst.gridgrind.excel.DataValidationFixtures.report(
                      new ResolvedDataValidation(
                          new DataValidationRuleSpec.ExplicitList(List.of()),
                          true,
                          false,
                          Optional.empty(),
                          Optional.empty())))),
          snapshots);
      assertEquals(
          "\"\"",
          sheet.getCTWorksheet().getDataValidations().getDataValidationArray(0).getFormula1());
      assertEquals(
          List.of(AnalysisFindingCode.DATA_VALIDATION_EMPTY_EXPLICIT_LIST),
          controller.dataValidationHealthFindings("Budget", sheet).stream()
              .map(AnalysisFindingReport::code)
              .distinct()
              .toList());
    }
  }

  private static ResolvedDataValidation explicitListValidation() {
    return new ResolvedDataValidation(
        new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
        true,
        false,
        Optional.of(new ResolvedValidationPrompt("Status", "Choose one workflow state.", true)),
        Optional.of(
            new ResolvedValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP,
                "Invalid status",
                "Use one of the allowed values.",
                true)));
  }

  private static ResolvedValidationErrorAlert warningAlert() {
    return new ResolvedValidationErrorAlert(
        ExcelDataValidationErrorStyle.WARNING,
        "Outside range",
        "Use a value greater than 0.5.",
        true);
  }

  private DataValidationEntryReport invokeToSnapshot(
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

    DataValidationEntryReport.Supported snapshot =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class,
            ExcelDataValidationController.toSnapshot(validation, List.of("Z1")));

    DataValidationRuleSpec rule = snapshot.validation().rule();
    ExcelComparisonOperator actualOperator =
        switch (rule) {
          case DataValidationRuleSpec.WholeNumber wholeNumber -> wholeNumber.operator();
          case DataValidationRuleSpec.DecimalNumber decimalNumber -> decimalNumber.operator();
          case DataValidationRuleSpec.DateRule dateRule -> dateRule.operator();
          case DataValidationRuleSpec.TimeRule timeRule -> timeRule.operator();
          case DataValidationRuleSpec.TextLength textLength -> textLength.operator();
          default -> throw new AssertionError("unexpected rule type: " + rule);
        };
    assertEquals(expectedOperator, actualOperator);
  }

  private static ResolvedDataValidation expectedRoundTripDefinition(
      ResolvedDataValidation definition) {
    boolean suppressDropDownArrow = true;
    if (definition.rule() instanceof DataValidationRuleSpec.ExplicitList) {
      suppressDropDownArrow = definition.suppressDropDownArrow();
    }
    return new ResolvedDataValidation(
        definition.rule(),
        definition.allowBlank(),
        suppressDropDownArrow,
        definition.prompt(),
        definition.errorAlert());
  }

  private void assertRoundTrip(String range, ResolvedDataValidation definition) throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      controller.setDataValidation(sheet, range, definition);

      List<DataValidationEntryReport> snapshots = controller.dataValidations(sheet);

      assertEquals(
          List.of(
              new DataValidationEntryReport.Supported(
                  List.of(range).stream()
                      .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                      .toList(),
                  dev.erst.gridgrind.excel.DataValidationFixtures.report(
                      expectedRoundTripDefinition(definition)))),
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
    @SuppressWarnings("PMD.ReturnEmptyCollectionRatherThanNull")
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

  private boolean isBrokenFormulaFinding(AnalysisFindingReport finding) {
    return finding.code() == AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA;
  }

  private boolean isOverlapFinding(AnalysisFindingReport finding) {
    return finding.code() == AnalysisFindingCode.DATA_VALIDATION_OVERLAPPING_RULES;
  }
}
