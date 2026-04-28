package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.contract.dto.FormulaUdfFunctionInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfToolpackInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Formula-environment request-executor regressions for Phase 4 parity. */
class FormulaEnvironmentRequestExecutorTest {
  @Test
  void evaluatesExternalWorkbookReferencesThroughFormulaEnvironment() throws Exception {
    ExternalFormulaScenario scenario = createExternalFormulaScenario(true);

    GridGrindResponse.Success success =
        assertInstanceOf(
            GridGrindResponse.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(
                            scenario.workbookPath().toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        new FormulaEnvironmentInput(
                            List.of(
                                new FormulaExternalWorkbookInput(
                                    "referenced.xlsx",
                                    scenario.referencedWorkbookPath().toString())),
                            FormulaMissingWorkbookPolicy.ERROR,
                            List.of()),
                        List.of(),
                        List.of(
                            inspect(
                                "cells",
                                new CellSelector.ByAddresses("Ops", List.of("B1")),
                                new InspectionQuery.GetCells())))));

    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formula =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class,
            ((InspectionResult.CellsResult) success.inspections().getFirst()).cells().getFirst());
    assertEquals(
        7.5d,
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class, formula.evaluation())
            .numberValue());
  }

  @Test
  void usesCachedFormulaValueWhenMissingExternalWorkbookPolicyAllowsIt() throws Exception {
    ExternalFormulaScenario scenario = createExternalFormulaScenario(true);

    GridGrindResponse.Success success =
        assertInstanceOf(
            GridGrindResponse.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(
                            scenario.workbookPath().toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        new FormulaEnvironmentInput(
                            List.of(), FormulaMissingWorkbookPolicy.USE_CACHED_VALUE, List.of()),
                        List.of(),
                        List.of(
                            inspect(
                                "cells",
                                new CellSelector.ByAddresses("Ops", List.of("B1")),
                                new InspectionQuery.GetCells())))));

    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formula =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class,
            ((InspectionResult.CellsResult) success.inspections().getFirst()).cells().getFirst());
    assertEquals(
        7.5d,
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class, formula.evaluation())
            .numberValue());
  }

  @Test
  void evaluatesRegisteredTemplateBackedUserDefinedFunctions() throws Exception {
    Path workbookPath = createUdfFormulaWorkbook();

    GridGrindResponse.Success success =
        assertInstanceOf(
            GridGrindResponse.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        new FormulaEnvironmentInput(
                            List.of(),
                            FormulaMissingWorkbookPolicy.ERROR,
                            List.of(
                                new FormulaUdfToolpackInput(
                                    "math",
                                    List.of(
                                        new FormulaUdfFunctionInput(
                                            "DOUBLE", 1, null, "ARG1*2"))))),
                        List.of(),
                        List.of(
                            inspect(
                                "cells",
                                new CellSelector.ByAddresses("Ops", List.of("B1")),
                                new InspectionQuery.GetCells())))));

    dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formula =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.CellReport.FormulaReport.class,
            ((InspectionResult.CellsResult) success.inspections().getFirst()).cells().getFirst());
    assertEquals(
        42.0d,
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.NumberReport.class, formula.evaluation())
            .numberValue());
  }

  @Test
  void reportsUnregisteredUserDefinedFunctionsPrecisely() throws Exception {
    Path workbookPath = createUdfFormulaWorkbook();

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(calculateAll()),
                        null,
                        List.of(),
                        List.of())));

    assertEquals(GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION, failure.problem().code());
    assertEquals(
        java.util.Optional.of("DOUBLE(A1)"),
        DefaultGridGrindRequestExecutorTestSupport.calculationPreflightContext(failure).formula());
  }

  @Test
  void evaluatesOnlyRequestedFormulaCellsWhenTargetedOperationIsUsed() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-targeted-protocol-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success success =
        assertInstanceOf(
            GridGrindResponse.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        executionPolicy(
                            calculateTargets(new CellSelector.QualifiedAddress("Budget", "B1"))),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A1"),
                                new MutationAction.SetCell(new CellInput.Numeric(2.0d))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "B1"),
                                new MutationAction.SetCell(formulaCell("A1*2"))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "C1"),
                                new MutationAction.SetCell(formulaCell("A1*3"))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A1"),
                                new MutationAction.SetCell(new CellInput.Numeric(4.0d)))),
                        inspections())));

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(8.0d, cachedFormulaValue(workbookPath, "Budget", "B1"));
    assertEquals(0.0d, cachedFormulaValue(workbookPath, "Budget", "C1"));
  }

  @Test
  void clearFormulaCachesRemovesPersistedCachedResults() throws Exception {
    Path workbookPath = Files.createTempFile("gridgrind-cleared-formula-caches-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success seeded =
        assertInstanceOf(
            GridGrindResponse.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        executionPolicy(calculateAll()),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Budget"),
                                new MutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Budget", "A1"),
                                new MutationAction.SetCell(new CellInput.Numeric(2.0d))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "B1"),
                                new MutationAction.SetCell(formulaCell("A1*2"))),
                            mutate(
                                new CellSelector.ByAddress("Budget", "C1"),
                                new MutationAction.SetCell(formulaCell("A1*3")))),
                        inspections())));
    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(seeded));

    GridGrindResponse.Success success =
        assertInstanceOf(
            GridGrindResponse.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
                        executionPolicy(clearFormulaCaches()),
                        null,
                        List.of(),
                        List.of())));

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertNull(cachedFormulaRawValue(workbookPath, "Budget", "B1"));
    assertNull(cachedFormulaRawValue(workbookPath, "Budget", "C1"));
  }

  private static String savedPath(GridGrindResponse.Success success) {
    return switch (success.persistence()) {
      case GridGrindResponse.PersistenceOutcome.SavedAs savedAs -> savedAs.executionPath();
      case GridGrindResponse.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponse.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("Expected the workbook to be persisted");
    };
  }

  private static ExternalFormulaScenario createExternalFormulaScenario(boolean seedCachedValue)
      throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-protocol-external-formula-");
    Path referencedWorkbookPath = directory.resolve("referenced.xlsx");
    Path workbookPath = directory.resolve("external-formula.xlsx");

    try (XSSFWorkbook referencedWorkbook = new XSSFWorkbook()) {
      referencedWorkbook.createSheet("Rates").createRow(0).createCell(0).setCellValue(7.5d);
      try (OutputStream outputStream = Files.newOutputStream(referencedWorkbookPath)) {
        referencedWorkbook.write(outputStream);
      }
    }

    try (Workbook referencedWorkbook = WorkbookFactory.create(referencedWorkbookPath.toFile());
        XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.linkExternalWorkbook("referenced.xlsx", referencedWorkbook);
      workbook.setCellFormulaValidation(false);
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("External");
      sheet.getRow(0).createCell(1).setCellFormula("[referenced.xlsx]Rates!$A$1");
      if (seedCachedValue) {
        var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.setupReferencedWorkbooks(
            Map.of(
                "referenced.xlsx",
                referencedWorkbook.getCreationHelper().createFormulaEvaluator()));
        evaluator.evaluateFormulaCell(sheet.getRow(0).getCell(1));
      }
      try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    return new ExternalFormulaScenario(workbookPath, referencedWorkbookPath);
  }

  private static Path createUdfFormulaWorkbook() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-protocol-udf-", ".xlsx");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.setCellFormulaValidation(false);
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue(21.0d);
      sheet.getRow(0).createCell(1).setCellFormula("DOUBLE(A1)");
      try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    return workbookPath;
  }

  private static double cachedFormulaValue(Path workbookPath, String sheetName, String address)
      throws IOException {
    try (XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(workbookPath.toFile())) {
      var reference = new org.apache.poi.ss.util.CellReference(address);
      var cell =
          workbook.getSheet(sheetName).getRow(reference.getRow()).getCell(reference.getCol());
      assertEquals(CellType.FORMULA, cell.getCellType());
      assertEquals(CellType.NUMERIC, cell.getCachedFormulaResultType());
      return cell.getNumericCellValue();
    }
  }

  private static String cachedFormulaRawValue(Path workbookPath, String sheetName, String address)
      throws IOException {
    try (XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(workbookPath.toFile())) {
      var reference = new org.apache.poi.ss.util.CellReference(address);
      var cell =
          (org.apache.poi.xssf.usermodel.XSSFCell)
              workbook.getSheet(sheetName).getRow(reference.getRow()).getCell(reference.getCol());
      assertEquals(CellType.FORMULA, cell.getCellType());
      return cell.getCTCell().isSetV() ? cell.getCTCell().getV() : null;
    }
  }

  private record ExternalFormulaScenario(Path workbookPath, Path referencedWorkbookPath) {}
}
