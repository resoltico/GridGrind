package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationHealthReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.WorkbookFindingsReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
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
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      ResolvedDataValidation definition =
          new ResolvedDataValidation(
              new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
              true,
              false,
              Optional.of(new ResolvedValidationPrompt("Status", "Pick one workflow state.", true)),
              Optional.of(
                  new ResolvedValidationErrorAlert(
                      ExcelDataValidationErrorStyle.STOP,
                      "Invalid status",
                      "Use one of the allowed values.",
                      true)));
      budget.metadata().setDataValidation("A1:C3", definition);

      ExcelWorkbookIntrospector introspector = new ExcelWorkbookIntrospector();
      dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult beforeClear =
          cast(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetDataValidations(
                      new StepId("all"), "Budget", new ExcelRangeSelection.All())));
      dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult selected =
          cast(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetDataValidations(
                      new StepId("selected"),
                      "Budget",
                      new ExcelRangeSelection.Selected(List.of("B2")))));

      assertEquals(1, beforeClear.validations().size());
      DataValidationEntryReport.Supported supported =
          assertInstanceOf(
              DataValidationEntryReport.Supported.class, beforeClear.validations().getFirst());
      assertEquals(
          List.of("A1:C3"),
          supported.ranges().stream()
              .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
              .toList());
      assertEquals(
          dev.erst.gridgrind.excel.DataValidationFixtures.report(definition),
          supported.validation());
      assertEquals(beforeClear.validations(), selected.validations());

      budget.metadata().clearDataValidations(new ExcelRangeSelection.Selected(List.of("B2")));

      dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult afterClear =
          cast(
              dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult.class,
              introspector.execute(
                  workbook,
                  new WorkbookReadCommand.GetDataValidations(
                      new StepId("afterClear"), "Budget", new ExcelRangeSelection.All())));

      assertEquals(1, afterClear.validations().size());
      DataValidationEntryReport.Supported retained =
          assertInstanceOf(
              DataValidationEntryReport.Supported.class, afterClear.validations().getFirst());
      assertEquals(
          List.of("A1:C1", "A3:C3", "A2", "C2"),
          retained.ranges().stream()
              .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::value)
              .toList());
      assertEquals(
          dev.erst.gridgrind.excel.DataValidationFixtures.report(definition),
          retained.validation());
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

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      WorkbookAnalyzer analyzer = new WorkbookAnalyzer();

      DataValidationHealthReport health =
          analyzer.dataValidationHealth(workbook, new ExcelSheetSelection.All());
      dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.DataValidationHealthResult executed =
          cast(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.DataValidationHealthResult.class,
              analyzer.execute(
                  workbook,
                  new WorkbookLocation.StoredWorkbook(workbookPath),
                  new WorkbookReadCommand.AnalyzeDataValidationHealth(
                      new StepId("validationHealth"), new ExcelSheetSelection.All())));
      WorkbookFindingsReport workbookFindings = analyzer.workbookFindings(workbook);

      assertEquals(2, health.checkedValidationCount());
      assertEquals(2, executed.analysis().checkedValidationCount());
      assertTrue(
          health.findings().stream()
              .map(AnalysisFindingReport::code)
              .toList()
              .containsAll(
                  List.of(
                      AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
                      AnalysisFindingCode.DATA_VALIDATION_OVERLAPPING_RULES)));
      assertTrue(
          workbookFindings.findings().stream()
              .map(AnalysisFindingReport::code)
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
