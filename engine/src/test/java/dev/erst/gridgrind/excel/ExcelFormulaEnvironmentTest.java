package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.model.formula.FormulaCapabilityKind;
import dev.erst.gridgrind.core.model.formula.FormulaCellTarget;
import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.core.model.formula.FormulaUdfFunctionSpec;
import dev.erst.gridgrind.core.model.formula.FormulaUdfToolpackSpec;
import dev.erst.gridgrind.core.model.formula.ResolvedExternalWorkbook;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
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

/** Regression tests for the workbook formula environment and lifecycle controls. */
class ExcelFormulaEnvironmentTest {

  @Test
  void evaluatesExternalWorkbookBindingsAndPersistsCachedResults() throws Exception {
    ExternalFormulaScenario scenario = createExternalFormulaScenario(true);
    Path outputPath = XlsxRoundTrip.newWorkbookPath("gridgrind-external-formula-");

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            scenario.workbookPath(),
            new ResolvedFormulaEnvironment(
                List.of(
                    new ResolvedExternalWorkbook(
                        "referenced.xlsx", scenario.referencedWorkbookPath())),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      workbook.formulas().evaluateAll();
      workbook
          .persistence()
          .save(
              outputPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(7.5d, cachedFormulaValue(outputPath, "Ops", "B1"));
  }

  @Test
  void usesCachedFormulaResultsWhenMissingExternalWorkbookPolicyAllowsIt() throws Exception {
    ExternalFormulaScenario scenario = createExternalFormulaScenario(true);
    Path outputPath = XlsxRoundTrip.newWorkbookPath("gridgrind-external-cached-");

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            scenario.workbookPath(),
            new ResolvedFormulaEnvironment(
                List.of(), FormulaMissingWorkbookPolicy.USE_CACHED_VALUE, List.of()),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      workbook.formulas().evaluateAll();
      workbook
          .persistence()
          .save(
              outputPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(7.5d, cachedFormulaValue(outputPath, "Ops", "B1"));
  }

  @Test
  void evaluatesTemplateBackedUserDefinedFunctions() throws Exception {
    Path workbookPath = createUdfFormulaWorkbook();
    Path outputPath = XlsxRoundTrip.newWorkbookPath("gridgrind-udf-formula-");

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            workbookPath,
            new ResolvedFormulaEnvironment(
                List.of(),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of(
                    new FormulaUdfToolpackSpec(
                        "math", List.of(new FormulaUdfFunctionSpec("DOUBLE", 1, 1, "ARG1*2"))))),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      workbook.formulas().evaluateAll();
      workbook
          .persistence()
          .save(
              outputPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(42.0d, cachedFormulaValue(outputPath, "Ops", "B1"));
  }

  @Test
  void targetedFormulaEvaluationRefreshesOnlyRequestedCachedResults() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-targeted-formula-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", new CellValue.NumberValue(2.0d));
      workbook.sheet("Budget").cells().setCell("B1", new ResolvedFormula("A1*2"));
      workbook.sheet("Budget").cells().setCell("C1", new ResolvedFormula("A1*3"));
      workbook.formulas().evaluateAll();
      workbook.sheet("Budget").cells().setCell("A1", new CellValue.NumberValue(4.0d));
      workbook
          .formulas()
          .evaluate(List.of(new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1"))));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(8.0d, cachedFormulaValue(workbookPath, "Budget", "B1"));
    assertEquals(6.0d, cachedFormulaValue(workbookPath, "Budget", "C1"));
  }

  @Test
  void clearFormulaCachesRemovesPersistedCachedResults() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-cleared-formula-caches-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", new CellValue.NumberValue(2.0d));
      workbook.sheet("Budget").cells().setCell("B1", new ResolvedFormula("A1*2"));
      workbook.sheet("Budget").cells().setCell("C1", new ResolvedFormula("A1*3"));
      workbook.formulas().evaluateAll();
      workbook.formulas().clearCaches();
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertNull(cachedFormulaRawValue(workbookPath, "Budget", "B1"));
    assertNull(cachedFormulaRawValue(workbookPath, "Budget", "C1"));
  }

  @Test
  void assessAllFormulaCapabilitiesClassifiesRuntimeOutcomes() throws Exception {
    ExternalFormulaScenario externalScenario = createExternalFormulaScenario(false);
    Path udfWorkbookPath = createUdfFormulaWorkbook();
    Path unsupportedWorkbookPath = createUnsupportedFormulaWorkbook();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", new CellValue.NumberValue(2.0d));
      workbook.sheet("Budget").cells().setCell("B1", new ResolvedFormula("A1*2"));

      assertEquals(
          List.of(
              new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
                  "A1*2",
                  FormulaCapabilityKind.EVALUABLE_NOW,
                  java.util.Optional.empty(),
                  java.util.Optional.empty())),
          workbook.formulas().assessAllCapabilities());
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            externalScenario.workbookPath(), ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      CalculationReport.FormulaCapability assessment =
          workbook.formulas().assessAllCapabilities().getFirst();

      assertEquals(FormulaCapabilityKind.UNEVALUABLE_NOW, assessment.capability());
      assertEquals(
          GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK, assessment.problemCode().orElse(null));
      assertTrue(assessment.message().orElseThrow().contains("Missing external workbook"));
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(udfWorkbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      CalculationReport.FormulaCapability assessment =
          workbook.formulas().assessAllCapabilities().getFirst();

      assertEquals(FormulaCapabilityKind.UNEVALUABLE_NOW, assessment.capability());
      assertEquals(
          GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION,
          assessment.problemCode().orElse(null));
      assertTrue(assessment.message().orElseThrow().contains("DOUBLE"));
    }

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            unsupportedWorkbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      CalculationReport.FormulaCapability assessment =
          workbook.formulas().assessAllCapabilities().getFirst();

      assertEquals(FormulaCapabilityKind.UNEVALUABLE_NOW, assessment.capability());
      assertEquals(GridGrindProblemCode.UNSUPPORTED_FORMULA, assessment.problemCode().orElse(null));
      assertTrue(assessment.message().orElseThrow().contains("APP.TITLE"));
    }
  }

  @Test
  void assessFormulaCellCapabilitiesSupportsExplicitTargetsAndRejectsInvalidTargets()
      throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", new CellValue.NumberValue(2.0d));
      workbook.sheet("Budget").cells().setCell("B1", new ResolvedFormula("A1*2"));

      assertEquals(
          List.of(
              new CalculationReport.FormulaCapability(
                  new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")),
                  "A1*2",
                  FormulaCapabilityKind.EVALUABLE_NOW,
                  java.util.Optional.empty(),
                  java.util.Optional.empty())),
          workbook
              .formulas()
              .assessCapabilities(
                  List.of(new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1")))));
      assertThrows(NullPointerException.class, () -> workbook.formulas().assessCapabilities(null));
      assertThrows(
          NullPointerException.class,
          () -> workbook.formulas().assessCapabilities(List.of((FormulaCellTarget) null)));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              workbook
                  .formulas()
                  .assessCapabilities(
                      List.of(
                          new FormulaCellTarget(new SheetName("Budget"), new CellAddress("A1")))));
      assertThrows(
          dev.erst.gridgrind.core.problem.InvalidRequestException.class,
          () ->
              workbook
                  .formulas()
                  .assessCapabilities(
                      List.of(
                          new FormulaCellTarget(new SheetName("Budget"), new CellAddress(":")))));
      assertThrows(
          CellNotFoundException.class,
          () ->
              workbook
                  .formulas()
                  .assessCapabilities(
                      List.of(
                          new FormulaCellTarget(new SheetName("Budget"), new CellAddress("Z99")))));
    }

    Path invalidWorkbookPath = createInvalidFormulaWorkbook();
    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            invalidWorkbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      CalculationReport.FormulaCapability assessment =
          workbook
              .formulas()
              .assessCapabilities(
                  List.of(new FormulaCellTarget(new SheetName("Budget"), new CellAddress("B1"))))
              .getFirst();

      assertEquals(FormulaCapabilityKind.UNPARSEABLE_BY_POI, assessment.capability());
      assertEquals(GridGrindProblemCode.INVALID_FORMULA, assessment.problemCode().orElse(null));
      assertEquals("Budget", assessment.cell().sheetName().value());
      assertEquals("B1", assessment.cell().address().value());
      assertEquals("SUM(", assessment.formula());
      assertTrue(assessment.message().orElseThrow().contains("Invalid formula at Budget!B1"));
    }
  }

  @Test
  void assessAllFormulaCapabilitiesRethrowsUnexpectedRuntimeFailures() throws Exception {
    try (XSSFWorkbook poiWorkbook = new XSSFWorkbook();
        ExcelWorkbook workbook =
            new ExcelWorkbook(
                poiWorkbook,
                FormulaRuntimeTestDouble.alwaysFail(new IllegalStateException("boom")))) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").cells().setCell("A1", new ResolvedFormula("1+1"));

      IllegalStateException failure =
          assertThrows(IllegalStateException.class, workbook.formulas()::assessAllCapabilities);

      assertEquals("boom", failure.getMessage());
    }
  }

  private static ExternalFormulaScenario createExternalFormulaScenario(boolean seedCachedValue)
      throws IOException {
    Path directory =
        ExcelTempFiles.createManagedTempDirectory("gridgrind-external-formula-scenario-");
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
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-udf-workbook-");

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
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-unsupported-formula-");

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
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-invalid-formula-");

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
