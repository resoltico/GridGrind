package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.calculateAll;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.calculateAllAndMarkRecalculateOnOpen;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.calculateTargets;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.clearFormulaCaches;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.markRecalculateOnOpen;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.CalculationExecutionStatus;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.FormulaCapabilityKind;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelFormulaCapabilityAssessment;
import dev.erst.gridgrind.excel.ExcelFormulaCapabilityIssue;
import dev.erst.gridgrind.excel.ExcelFormulaCapabilityKind;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Direct unit coverage for Phase 6 calculation-policy execution helpers. */
class CalculationPolicyExecutorTest {
  @Test
  void normalizesPoliciesAndExposesModeGuards() {
    CalculationPolicyInput normalized = CalculationPolicyExecutor.normalize(null);

    assertTrue(normalized.isDefault());
    assertTrue(CalculationPolicyExecutor.allowsEventRead(null));
    assertFalse(CalculationPolicyExecutor.allowsEventRead(calculateAll()));
    assertTrue(CalculationPolicyExecutor.allowsStreamingWrite(markRecalculateOnOpen()));
    assertFalse(CalculationPolicyExecutor.allowsEventRead(markRecalculateOnOpen()));
    assertFalse(CalculationPolicyExecutor.allowsStreamingWrite(calculateAll()));
    assertFalse(CalculationPolicyExecutor.requiresMutationPrefix(null));
    assertTrue(CalculationPolicyExecutor.requiresMutationPrefix(clearFormulaCaches()));
    assertTrue(
        CalculationPolicyExecutor.requiresMutationPrefix(
            calculateTargets(new CellSelector.QualifiedAddress("Budget", "B1"))));

    CalculationReport report =
        CalculationPolicyExecutor.notRequestedReport(markRecalculateOnOpen());
    assertTrue(report.policy().markRecalculateOnOpen());
    assertEquals(CalculationExecutionStatus.NOT_REQUESTED, report.execution().status());
    assertNull(report.preflight());
  }

  @Test
  void preflightAndExecutionSucceedForSupportedPolicies() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      CalculationPolicyExecutor.PreflightOutcome notRequested =
          CalculationPolicyExecutor.preflight(workbook, null);
      CalculationPolicyExecutor.PreflightOutcome clearCachesOnly =
          CalculationPolicyExecutor.preflight(workbook, clearFormulaCaches());

      assertNull(notRequested.report());
      assertEquals(0, notRequested.evaluationTargetCount());
      assertNull(clearCachesOnly.report());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      CalculationPolicyExecutor.ExecutionOutcome doNotCalculate =
          CalculationPolicyExecutor.execute(workbook, new CalculationPolicyInput(null, false), 0);

      assertEquals(CalculationExecutionStatus.NOT_REQUESTED, doNotCalculate.report().status());
      assertFalse(doNotCalculate.report().markRecalculateOnOpenApplied());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(2.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));
      workbook.sheet("Budget").setCell("C1", ExcelCellValue.formula("A1*3"));

      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, calculateAll());
      CalculationPolicyExecutor.ExecutionOutcome execution =
          CalculationPolicyExecutor.execute(
              workbook, calculateAllAndMarkRecalculateOnOpen(), preflight.evaluationTargetCount());

      assertNull(preflight.failure());
      assertEquals(CalculationReport.Scope.WORKBOOK, preflight.report().scope());
      assertEquals(2, preflight.evaluationTargetCount());
      assertNull(execution.failure());
      assertEquals(CalculationExecutionStatus.SUCCEEDED, execution.report().status());
      assertEquals(2, execution.report().evaluatedFormulaCount());
      assertTrue(execution.report().markRecalculateOnOpenApplied());
      assertTrue(workbook.formulas().recalculateOnOpenEnabled());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(5.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));

      CalculationPolicyInput targeted =
          calculateTargets(new CellSelector.QualifiedAddress("Budget", "B1"));
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, targeted);
      CalculationPolicyExecutor.ExecutionOutcome execution =
          CalculationPolicyExecutor.execute(workbook, targeted, preflight.evaluationTargetCount());

      assertNull(preflight.failure());
      assertEquals(CalculationReport.Scope.TARGETS, preflight.report().scope());
      assertEquals(1, execution.report().evaluatedFormulaCount());
      assertFalse(execution.report().markRecalculateOnOpenApplied());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(2.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));
      workbook.formulas().evaluateAll();

      CalculationPolicyExecutor.ExecutionOutcome clearCaches =
          CalculationPolicyExecutor.execute(
              workbook,
              new CalculationPolicyInput(new CalculationStrategyInput.ClearCachesOnly(), true),
              0);
      CalculationPolicyExecutor.ExecutionOutcome markOnly =
          CalculationPolicyExecutor.execute(workbook, markRecalculateOnOpen(), 0);

      assertTrue(clearCaches.report().cachesCleared());
      assertTrue(clearCaches.report().markRecalculateOnOpenApplied());
      assertEquals(CalculationExecutionStatus.SUCCEEDED, markOnly.report().status());
      assertTrue(markOnly.report().markRecalculateOnOpenApplied());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(5.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));

      CalculationPolicyInput targetedAndMarked =
          new CalculationPolicyInput(
              new CalculationStrategyInput.EvaluateTargets(
                  List.of(new CellSelector.QualifiedAddress("Budget", "B1"))),
              true);
      CalculationPolicyExecutor.ExecutionOutcome execution =
          CalculationPolicyExecutor.execute(workbook, targetedAndMarked, 1);

      assertEquals(CalculationExecutionStatus.SUCCEEDED, execution.report().status());
      assertTrue(execution.report().markRecalculateOnOpenApplied());
    }
  }

  @Test
  void preflightAndExecutionFailuresCarryClassificationAndValidation() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.open(createMixedFailureWorkbook())) {
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, calculateAll());

      assertEquals(GridGrindProblemCode.INVALID_FORMULA, preflight.failure().code());
      assertEquals(1, preflight.report().summary().evaluableNowCount());
      assertEquals(1, preflight.report().summary().unevaluableNowCount());
      assertEquals(1, preflight.report().summary().unparseableByPoiCount());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(createInvalidFormulaWorkbook())) {
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, calculateAll());
      CalculationPolicyExecutor.ExecutionOutcome targetedExecution =
          CalculationPolicyExecutor.execute(
              workbook, calculateTargets(new CellSelector.QualifiedAddress("Budget", "B1")), 1);

      assertEquals(GridGrindProblemCode.INVALID_FORMULA, preflight.failure().code());
      assertEquals(CalculationPolicyExecutor.Phase.PREFLIGHT, preflight.failure().phase());
      assertEquals("Budget", preflight.failure().sheetName());
      assertEquals("B1", preflight.failure().address());
      assertEquals("SUM(", preflight.failure().formula());
      assertEquals(1, preflight.report().summary().unparseableByPoiCount());
      assertEquals(CalculationExecutionStatus.FAILED, targetedExecution.report().status());
      assertEquals(CalculationPolicyExecutor.Phase.EXECUTION, targetedExecution.failure().phase());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(createMissingExternalWorkbook())) {
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, calculateAll());

      assertEquals(GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK, preflight.failure().code());
      assertEquals(1, preflight.report().summary().unevaluableNowCount());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(createUdfFormulaWorkbook())) {
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, calculateAll());

      assertEquals(
          GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION, preflight.failure().code());
      assertTrue(preflight.failure().message().contains("DOUBLE"));
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(createUnsupportedFormulaWorkbook())) {
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, calculateAll());
      CalculationPolicyExecutor.ExecutionOutcome execution =
          CalculationPolicyExecutor.execute(workbook, calculateAll(), 1);

      assertEquals(GridGrindProblemCode.UNSUPPORTED_FORMULA, preflight.failure().code());
      assertEquals(CalculationExecutionStatus.FAILED, execution.report().status());
      assertEquals(CalculationPolicyExecutor.Phase.EXECUTION, execution.failure().phase());
      assertNotNull(execution.failure().exception());
      assertTrue(execution.report().message().contains("APP.TITLE"));
    }

    try (ExcelWorkbook workbook = failingClearCachesWorkbook()) {
      CalculationPolicyExecutor.ExecutionOutcome clearCachesFailure =
          CalculationPolicyExecutor.execute(workbook, clearFormulaCaches(), 0);

      assertEquals(CalculationExecutionStatus.FAILED, clearCachesFailure.report().status());
      assertEquals(CalculationPolicyExecutor.Phase.EXECUTION, clearCachesFailure.failure().phase());
      assertTrue(clearCachesFailure.report().message().contains("clear caches failed"));
    }

    assertEquals(
        "evaluationTargetCount must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CalculationPolicyExecutor.PreflightOutcome(null, -1, null))
            .getMessage());
    assertEquals(
        "report must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new CalculationPolicyExecutor.ExecutionOutcome(null, null))
            .getMessage());
    assertEquals(
        "phase must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new CalculationPolicyExecutor.FailureDetail(
                        GridGrindProblemCode.INVALID_FORMULA,
                        null,
                        "Budget",
                        "B1",
                        "SUM(",
                        "bad",
                        null))
            .getMessage());
    assertEquals(
        "code or exception must be present",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CalculationPolicyExecutor.FailureDetail(
                        null,
                        CalculationPolicyExecutor.Phase.PREFLIGHT,
                        null,
                        null,
                        null,
                        "bad",
                        null))
            .getMessage());
    assertEquals(
        "message must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CalculationPolicyExecutor.FailureDetail(
                        GridGrindProblemCode.INVALID_FORMULA,
                        CalculationPolicyExecutor.Phase.PREFLIGHT,
                        null,
                        null,
                        null,
                        " ",
                        null))
            .getMessage());
    assertInstanceOf(
        RuntimeException.class,
        new CalculationPolicyExecutor.FailureDetail(
                CalculationPolicyExecutor.Phase.EXECUTION, new RuntimeException("boom"))
            .exception());
  }

  @Test
  void helperMappingsCoverAllCapabilityKindsAndProblemCodes() {
    ExcelFormulaCapabilityAssessment evaluable =
        new ExcelFormulaCapabilityAssessment(
            "Budget", "A1", "1+1", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null);
    ExcelFormulaCapabilityAssessment missingExternal =
        new ExcelFormulaCapabilityAssessment(
            "Budget",
            "B1",
            "[Rates.xlsx]Sheet1!A1",
            ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
            ExcelFormulaCapabilityIssue.MISSING_EXTERNAL_WORKBOOK,
            "missing");
    ExcelFormulaCapabilityAssessment unregistered =
        new ExcelFormulaCapabilityAssessment(
            "Budget",
            "C1",
            "DOUBLE(A1)",
            ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
            ExcelFormulaCapabilityIssue.UNREGISTERED_USER_DEFINED_FUNCTION,
            "unregistered");
    ExcelFormulaCapabilityAssessment unsupported =
        new ExcelFormulaCapabilityAssessment(
            "Budget",
            "D1",
            "APP.TITLE()",
            ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
            ExcelFormulaCapabilityIssue.UNSUPPORTED_FORMULA,
            "unsupported");
    ExcelFormulaCapabilityAssessment invalid =
        new ExcelFormulaCapabilityAssessment(
            "Budget",
            "E1",
            "SUM(",
            ExcelFormulaCapabilityKind.UNPARSEABLE_BY_POI,
            ExcelFormulaCapabilityIssue.INVALID_FORMULA,
            "invalid");

    CalculationReport.Summary summary =
        CalculationCapabilityMappings.summaryFor(List.of(evaluable, missingExternal, invalid));
    assertEquals(1, summary.evaluableNowCount());
    assertEquals(1, summary.unevaluableNowCount());
    assertEquals(1, summary.unparseableByPoiCount());
    assertEquals(null, CalculationCapabilityMappings.problemCodeFor(evaluable));
    assertEquals(
        GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK,
        CalculationCapabilityMappings.problemCodeFor(missingExternal));
    assertEquals(
        GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION,
        CalculationCapabilityMappings.problemCodeFor(unregistered));
    assertEquals(
        GridGrindProblemCode.UNSUPPORTED_FORMULA,
        CalculationCapabilityMappings.problemCodeFor(unsupported));
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        CalculationCapabilityMappings.problemCodeFor(invalid));
    assertEquals(
        FormulaCapabilityKind.EVALUABLE_NOW,
        CalculationCapabilityMappings.capabilityKindFor(ExcelFormulaCapabilityKind.EVALUABLE_NOW));
    assertEquals(
        FormulaCapabilityKind.UNEVALUABLE_NOW,
        CalculationCapabilityMappings.capabilityKindFor(
            ExcelFormulaCapabilityKind.UNEVALUABLE_NOW));
    assertEquals(
        FormulaCapabilityKind.UNPARSEABLE_BY_POI,
        CalculationCapabilityMappings.capabilityKindFor(
            ExcelFormulaCapabilityKind.UNPARSEABLE_BY_POI));
    assertEquals(0, CalculationCapabilityMappings.severityRank(invalid));
    assertEquals(1, CalculationCapabilityMappings.severityRank(missingExternal));
    assertEquals(2, CalculationCapabilityMappings.severityRank(unregistered));
    assertEquals(3, CalculationCapabilityMappings.severityRank(unsupported));
    assertEquals(
        "assessment.issue must not be null",
        assertThrows(
                NullPointerException.class,
                () -> CalculationCapabilityMappings.severityRank(evaluable))
            .getMessage());
  }

  private static Path createMissingExternalWorkbook() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-calculation-external-");
    Path referencedWorkbookPath = directory.resolve("referenced.xlsx");
    Path workbookPath = directory.resolve("external.xlsx");

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
      try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    return workbookPath;
  }

  private static Path createMixedFailureWorkbook() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-calculation-mixed-", ".xlsx");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.setCellFormulaValidation(false);
      XSSFSheet sheet = workbook.createSheet("Budget");
      sheet.createRow(0).createCell(0).setCellValue(2.0d);
      sheet.getRow(0).createCell(1).setCellFormula("A1*2");
      sheet.getRow(0).createCell(2).setCellFormula("APP.TITLE()");
      sheet.getRow(0).createCell(3).setCellFormula("A1*4");
      try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (FileSystem archive = FileSystems.newFileSystem(workbookPath, (ClassLoader) null)) {
      Path sheetXml = archive.getPath("/xl/worksheets/sheet1.xml");
      String xml = Files.readString(sheetXml, StandardCharsets.UTF_8);
      Files.writeString(
          sheetXml, xml.replace("<f>A1*2</f>", "<f>SUM(</f>"), StandardCharsets.UTF_8);
    }

    return workbookPath;
  }

  private static ExcelWorkbook failingClearCachesWorkbook() {
    return instantiateWorkbook(new IteratorFailingWorkbook());
  }

  private static ExcelWorkbook instantiateWorkbook(XSSFWorkbook workbook) {
    return ExcelWorkbook.wrap(workbook);
  }

  /** Workbook whose sheet iteration path fails so cache clearing can surface execution errors. */
  private static final class IteratorFailingWorkbook extends XSSFWorkbook {
    @Override
    public Iterator<Sheet> iterator() {
      throw new IllegalStateException("clear caches failed");
    }
  }

  private static Path createUdfFormulaWorkbook() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-calculation-udf-", ".xlsx");

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

  private static Path createUnsupportedFormulaWorkbook() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-calculation-unsupported-", ".xlsx");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.setCellFormulaValidation(false);
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellFormula("APP.TITLE()");
      try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    return workbookPath;
  }

  private static Path createInvalidFormulaWorkbook() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-calculation-invalid-", ".xlsx");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      sheet.createRow(0).createCell(0).setCellValue(2.0d);
      sheet.getRow(0).createCell(1).setCellFormula("A1*2");
      try (OutputStream outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    try (FileSystem archive = FileSystems.newFileSystem(workbookPath, (ClassLoader) null)) {
      Path sheetXml = archive.getPath("/xl/worksheets/sheet1.xml");
      String xml = Files.readString(sheetXml, StandardCharsets.UTF_8);
      Files.writeString(
          sheetXml, xml.replace("<f>A1*2</f>", "<f>SUM(</f>"), StandardCharsets.UTF_8);
    }

    return workbookPath;
  }
}
