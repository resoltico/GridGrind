package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Regression tests for the workbook formula environment and lifecycle controls. */
class ExcelFormulaEnvironmentTest {
  @Test
  void formulaEnvironmentDomainModelNormalizesAndValidatesInputs() {
    Path externalWorkbookPath = Path.of("tmp", "rates.xlsx");
    List<ExcelFormulaExternalWorkbookBinding> externalWorkbooks =
        new ArrayList<>(
            List.of(new ExcelFormulaExternalWorkbookBinding("Rates.xlsx", externalWorkbookPath)));
    List<ExcelFormulaUdfToolpack> udfToolpacks =
        new ArrayList<>(
            List.of(
                new ExcelFormulaUdfToolpack(
                    "math",
                    List.of(new ExcelFormulaUdfFunction("DOUBLE", 1, 2, " = ARG1 + ARG2 ")))));

    ExcelFormulaEnvironment environment =
        new ExcelFormulaEnvironment(externalWorkbooks, null, udfToolpacks);
    externalWorkbooks.clear();
    udfToolpacks.clear();

    assertFalse(environment.isDefault());
    assertEquals(ExcelFormulaMissingWorkbookPolicy.ERROR, environment.missingWorkbookPolicy());
    assertEquals(
        "ARG1 + ARG2",
        environment.udfToolpacks().getFirst().functions().getFirst().formulaTemplate());
    assertThrows(
        UnsupportedOperationException.class,
        () ->
            environment
                .externalWorkbooks()
                .add(new ExcelFormulaExternalWorkbookBinding("Other.xlsx", externalWorkbookPath)));

    ExcelFormulaRuntimeContext runtimeContext = environment.runtimeContext();
    assertTrue(runtimeContext.hasExternalWorkbookBinding("RATES.XLSX"));
    assertTrue(runtimeContext.hasUserDefinedFunction("double"));
    assertThrows(NullPointerException.class, () -> runtimeContext.hasExternalWorkbookBinding(null));
    assertThrows(NullPointerException.class, () -> runtimeContext.hasUserDefinedFunction(null));
    assertTrue(new ExcelFormulaEnvironment(null, null, null).isDefault());
    assertFalse(
        new ExcelFormulaEnvironment(null, ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE, null)
            .isDefault());
    assertFalse(
        new ExcelFormulaEnvironment(
                null,
                null,
                List.of(
                    new ExcelFormulaUdfToolpack(
                        "math", List.of(new ExcelFormulaUdfFunction("DOUBLE", 1, 1, "ARG1")))))
            .isDefault());
    assertEquals(
        Set.of(),
        new ExcelFormulaEnvironment(null, null, null).runtimeContext().externalWorkbookNames());
  }

  @Test
  void formulaEnvironmentDomainModelRejectsInvalidInputs() {
    Path externalWorkbookPath = Path.of("tmp", "rates.xlsx");

    assertThrows(
        NullPointerException.class,
        () -> new ExcelFormulaExternalWorkbookBinding(null, externalWorkbookPath));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelFormulaExternalWorkbookBinding(" ", externalWorkbookPath));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelFormulaExternalWorkbookBinding("Rates.xlsx", null));

    assertThrows(NullPointerException.class, () -> new ExcelFormulaUdfFunction(null, 1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelFormulaUdfFunction(" ", 1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelFormulaUdfFunction("1BAD", 1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelFormulaUdfFunction("DOUBLE", -1, 1, "ARG1"));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelFormulaUdfFunction("DOUBLE", 2, 1, "ARG1"));
    assertThrows(
        NullPointerException.class, () -> new ExcelFormulaUdfFunction("DOUBLE", 1, 1, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelFormulaUdfFunction("DOUBLE", 1, 1, " "));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelFormulaUdfFunction("DOUBLE", 1, 1, "ARG2"));

    assertThrows(NullPointerException.class, () -> new ExcelFormulaUdfToolpack(null, List.of()));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFormulaUdfToolpack(" ", List.of()));
    assertThrows(NullPointerException.class, () -> new ExcelFormulaUdfToolpack("math", null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelFormulaUdfToolpack("math", List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelFormulaUdfToolpack("math", List.of((ExcelFormulaUdfFunction) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaUdfToolpack(
                "math",
                List.of(
                    new ExcelFormulaUdfFunction("DOUBLE", 1, 1, "ARG1"),
                    new ExcelFormulaUdfFunction("double", 1, 1, "ARG1"))));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaEnvironment(
                List.of(
                    new ExcelFormulaExternalWorkbookBinding("Rates.xlsx", externalWorkbookPath),
                    new ExcelFormulaExternalWorkbookBinding("rates.xlsx", externalWorkbookPath)),
                ExcelFormulaMissingWorkbookPolicy.ERROR,
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelFormulaEnvironment(
                List.of((ExcelFormulaExternalWorkbookBinding) null),
                ExcelFormulaMissingWorkbookPolicy.ERROR,
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaEnvironment(
                List.of(),
                ExcelFormulaMissingWorkbookPolicy.ERROR,
                List.of(
                    new ExcelFormulaUdfToolpack(
                        "math", List.of(new ExcelFormulaUdfFunction("DOUBLE", 1, 1, "ARG1"))),
                    new ExcelFormulaUdfToolpack(
                        "Math", List.of(new ExcelFormulaUdfFunction("TRIPLE", 1, 1, "ARG1*3"))))));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelFormulaEnvironment(
                List.of(),
                ExcelFormulaMissingWorkbookPolicy.ERROR,
                List.of((ExcelFormulaUdfToolpack) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaEnvironment(
                List.of(),
                ExcelFormulaMissingWorkbookPolicy.ERROR,
                List.of(
                    new ExcelFormulaUdfToolpack(
                        "math", List.of(new ExcelFormulaUdfFunction("DOUBLE", 1, 1, "ARG1"))),
                    new ExcelFormulaUdfToolpack(
                        "stats", List.of(new ExcelFormulaUdfFunction("double", 1, 1, "ARG1"))))));

    ExcelFormulaUdfFunction dottedFunction = new ExcelFormulaUdfFunction("_pkg.DOUBLE2", 0, 0, "1");
    assertEquals("_pkg.DOUBLE2", dottedFunction.name());
  }

  @Test
  void evaluatesExternalWorkbookBindingsAndPersistsCachedResults() throws Exception {
    ExternalFormulaScenario scenario = createExternalFormulaScenario(true);
    Path outputPath = XlsxRoundTrip.newWorkbookPath("gridgrind-external-formula-");

    try (ExcelWorkbook workbook =
        ExcelWorkbook.open(
            scenario.workbookPath(),
            new ExcelFormulaEnvironment(
                List.of(
                    new ExcelFormulaExternalWorkbookBinding(
                        "referenced.xlsx", scenario.referencedWorkbookPath())),
                ExcelFormulaMissingWorkbookPolicy.ERROR,
                List.of()))) {
      workbook.formulas().evaluateAll();
      workbook.save(outputPath);
    }

    assertEquals(7.5d, cachedFormulaValue(outputPath, "Ops", "B1"));
  }

  @Test
  void usesCachedFormulaResultsWhenMissingExternalWorkbookPolicyAllowsIt() throws Exception {
    ExternalFormulaScenario scenario = createExternalFormulaScenario(true);
    Path outputPath = XlsxRoundTrip.newWorkbookPath("gridgrind-external-cached-");

    try (ExcelWorkbook workbook =
        ExcelWorkbook.open(
            scenario.workbookPath(),
            new ExcelFormulaEnvironment(
                List.of(), ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE, List.of()))) {
      workbook.formulas().evaluateAll();
      workbook.save(outputPath);
    }

    assertEquals(7.5d, cachedFormulaValue(outputPath, "Ops", "B1"));
  }

  @Test
  void evaluatesTemplateBackedUserDefinedFunctions() throws Exception {
    Path workbookPath = createUdfFormulaWorkbook();
    Path outputPath = XlsxRoundTrip.newWorkbookPath("gridgrind-udf-formula-");

    try (ExcelWorkbook workbook =
        ExcelWorkbook.open(
            workbookPath,
            new ExcelFormulaEnvironment(
                List.of(),
                ExcelFormulaMissingWorkbookPolicy.ERROR,
                List.of(
                    new ExcelFormulaUdfToolpack(
                        "math",
                        List.of(new ExcelFormulaUdfFunction("DOUBLE", 1, 1, "ARG1*2"))))))) {
      workbook.formulas().evaluateAll();
      workbook.save(outputPath);
    }

    assertEquals(42.0d, cachedFormulaValue(outputPath, "Ops", "B1"));
  }

  @Test
  void targetedFormulaEvaluationRefreshesOnlyRequestedCachedResults() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-targeted-formula-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(2.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));
      workbook.sheet("Budget").setCell("C1", ExcelCellValue.formula("A1*3"));
      workbook.formulas().evaluateAll();
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(4.0d));
      workbook.formulas().evaluate(List.of(new ExcelFormulaCellTarget("Budget", "B1")));
      workbook.save(workbookPath);
    }

    assertEquals(8.0d, cachedFormulaValue(workbookPath, "Budget", "B1"));
    assertEquals(6.0d, cachedFormulaValue(workbookPath, "Budget", "C1"));
  }

  @Test
  void clearFormulaCachesRemovesPersistedCachedResults() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-cleared-formula-caches-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(2.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));
      workbook.sheet("Budget").setCell("C1", ExcelCellValue.formula("A1*3"));
      workbook.formulas().evaluateAll();
      workbook.formulas().clearCaches();
      workbook.save(workbookPath);
    }

    assertNull(cachedFormulaRawValue(workbookPath, "Budget", "B1"));
    assertNull(cachedFormulaRawValue(workbookPath, "Budget", "C1"));
  }

  @Test
  void assessAllFormulaCapabilitiesClassifiesRuntimeOutcomes() throws Exception {
    ExternalFormulaScenario externalScenario = createExternalFormulaScenario(false);
    Path udfWorkbookPath = createUdfFormulaWorkbook();
    Path unsupportedWorkbookPath = createUnsupportedFormulaWorkbook();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(2.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));

      assertEquals(
          List.of(
              new ExcelFormulaCapabilityAssessment(
                  "Budget", "B1", "A1*2", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null)),
          workbook.formulas().assessAllCapabilities());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(externalScenario.workbookPath())) {
      ExcelFormulaCapabilityAssessment assessment =
          workbook.formulas().assessAllCapabilities().getFirst();

      assertEquals(ExcelFormulaCapabilityKind.UNEVALUABLE_NOW, assessment.capability());
      assertEquals(ExcelFormulaCapabilityIssue.MISSING_EXTERNAL_WORKBOOK, assessment.issue());
      assertTrue(assessment.message().contains("Missing external workbook"));
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(udfWorkbookPath)) {
      ExcelFormulaCapabilityAssessment assessment =
          workbook.formulas().assessAllCapabilities().getFirst();

      assertEquals(ExcelFormulaCapabilityKind.UNEVALUABLE_NOW, assessment.capability());
      assertEquals(
          ExcelFormulaCapabilityIssue.UNREGISTERED_USER_DEFINED_FUNCTION, assessment.issue());
      assertTrue(assessment.message().contains("DOUBLE"));
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(unsupportedWorkbookPath)) {
      ExcelFormulaCapabilityAssessment assessment =
          workbook.formulas().assessAllCapabilities().getFirst();

      assertEquals(ExcelFormulaCapabilityKind.UNEVALUABLE_NOW, assessment.capability());
      assertEquals(ExcelFormulaCapabilityIssue.UNSUPPORTED_FORMULA, assessment.issue());
      assertTrue(assessment.message().contains("APP.TITLE"));
    }
  }

  @Test
  void assessFormulaCellCapabilitiesSupportsExplicitTargetsAndRejectsInvalidTargets()
      throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(2.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));

      assertEquals(
          List.of(
              new ExcelFormulaCapabilityAssessment(
                  "Budget", "B1", "A1*2", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null)),
          workbook
              .formulas()
              .assessCapabilities(List.of(new ExcelFormulaCellTarget("Budget", "B1"))));
      assertThrows(NullPointerException.class, () -> workbook.formulas().assessCapabilities(null));
      assertThrows(
          NullPointerException.class,
          () -> workbook.formulas().assessCapabilities(List.of((ExcelFormulaCellTarget) null)));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              workbook
                  .formulas()
                  .assessCapabilities(List.of(new ExcelFormulaCellTarget("Budget", "A1"))));
      assertThrows(
          InvalidCellAddressException.class,
          () ->
              workbook
                  .formulas()
                  .assessCapabilities(List.of(new ExcelFormulaCellTarget("Budget", ":"))));
      assertThrows(
          CellNotFoundException.class,
          () ->
              workbook
                  .formulas()
                  .assessCapabilities(List.of(new ExcelFormulaCellTarget("Budget", "Z99"))));
    }

    Path invalidWorkbookPath = createInvalidFormulaWorkbook();
    try (ExcelWorkbook workbook = ExcelWorkbook.open(invalidWorkbookPath)) {
      ExcelFormulaCapabilityAssessment assessment =
          workbook
              .formulas()
              .assessCapabilities(List.of(new ExcelFormulaCellTarget("Budget", "B1")))
              .getFirst();

      assertEquals(ExcelFormulaCapabilityKind.UNPARSEABLE_BY_POI, assessment.capability());
      assertEquals(ExcelFormulaCapabilityIssue.INVALID_FORMULA, assessment.issue());
      assertEquals("Budget", assessment.sheetName());
      assertEquals("B1", assessment.address());
      assertEquals("SUM(", assessment.formula());
      assertTrue(assessment.message().contains("Invalid formula at Budget!B1"));
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
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.formula("1+1"));

      IllegalStateException failure =
          assertThrows(IllegalStateException.class, workbook.formulas()::assessAllCapabilities);

      assertEquals("boom", failure.getMessage());
    }
  }

  private static ExternalFormulaScenario createExternalFormulaScenario(boolean seedCachedValue)
      throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-external-formula-scenario-");
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
