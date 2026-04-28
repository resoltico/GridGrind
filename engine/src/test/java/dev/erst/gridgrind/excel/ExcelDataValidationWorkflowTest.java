package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused workflow tests for data-validation authoring, introspection, and analysis. */
class ExcelDataValidationWorkflowTest {
  @Test
  void introspectionPreservesDefinitionsAndSelectedClearSplitsCoverage() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      ExcelDataValidationDefinition definition =
          new ExcelDataValidationDefinition(
              new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
              true,
              false,
              new ExcelDataValidationPrompt("Status", "Pick one workflow state.", true),
              new ExcelDataValidationErrorAlert(
                  ExcelDataValidationErrorStyle.STOP,
                  "Invalid status",
                  "Use one of the allowed values.",
                  true));
      budget.setDataValidation("A1:C3", definition);

      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      WorkbookReadResult.DataValidationsResult beforeClear =
          cast(
              WorkbookReadResult.DataValidationsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetDataValidations(
                      "all", "Budget", new ExcelRangeSelection.All())));
      WorkbookReadResult.DataValidationsResult selected =
          cast(
              WorkbookReadResult.DataValidationsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetDataValidations(
                      "selected", "Budget", new ExcelRangeSelection.Selected(List.of("B2")))));

      assertEquals(1, beforeClear.validations().size());
      ExcelDataValidationSnapshot.Supported supported =
          assertInstanceOf(
              ExcelDataValidationSnapshot.Supported.class, beforeClear.validations().getFirst());
      assertEquals(List.of("A1:C3"), supported.ranges());
      assertEquals(definition, supported.validation());
      assertEquals(beforeClear.validations(), selected.validations());

      budget.clearDataValidations(new ExcelRangeSelection.Selected(List.of("B2")));

      WorkbookReadResult.DataValidationsResult afterClear =
          cast(
              WorkbookReadResult.DataValidationsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetDataValidations(
                      "afterClear", "Budget", new ExcelRangeSelection.All())));

      assertEquals(1, afterClear.validations().size());
      ExcelDataValidationSnapshot.Supported retained =
          assertInstanceOf(
              ExcelDataValidationSnapshot.Supported.class, afterClear.validations().getFirst());
      assertEquals(List.of("A1:C1", "A3:C3", "A2", "C2"), retained.ranges());
      assertEquals(definition, retained.validation());
    }
  }

  @Test
  void healthAnalysisFindsBrokenFormulaAndOverlaps() throws IOException {
    Path workbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-data-validation-health-", ".xlsx");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet budget = workbook.createSheet("Budget");
      DataValidationHelper helper = budget.getDataValidationHelper();
      DataValidation explicitList =
          helper.createValidation(
              helper.createExplicitListConstraint(new String[] {"Queued", "Done"}),
              new CellRangeAddressList(0, 3, 0, 0));
      DataValidation brokenFormulaList =
          helper.createValidation(
              helper.createFormulaListConstraint("#REF!"), new CellRangeAddressList(2, 4, 0, 0));
      budget.addValidationData(explicitList);
      budget.addValidationData(brokenFormulaList);
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookAnalyzer analyzer = new WorkbookAnalyzer();

      WorkbookAnalysis.DataValidationHealth health =
          analyzer.dataValidationHealth(workbook, new ExcelSheetSelection.All());
      WorkbookReadResult.DataValidationHealthResult executed =
          cast(
              WorkbookReadResult.DataValidationHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeDataValidationHealth(
                      "validationHealth", new ExcelSheetSelection.All())));
      WorkbookAnalysis.WorkbookFindings workbookFindings = analyzer.workbookFindings(workbook);

      assertEquals(2, health.checkedValidationCount());
      assertEquals(2, executed.analysis().checkedValidationCount());
      assertTrue(
          health.findings().stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
                      AnalysisFindingCode.DATA_VALIDATION_OVERLAPPING_RULES)));
      assertTrue(
          workbookFindings.findings().stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList()
              .containsAll(
                  List.of(
                      AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
                      AnalysisFindingCode.DATA_VALIDATION_OVERLAPPING_RULES)));
    }
  }

  private static <T> T cast(Class<T> expectedType, Object value) {
    return expectedType.cast(value);
  }
}
