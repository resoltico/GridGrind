package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.GridGrindCli;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.RecipeAdvisory;
import dev.erst.gridgrind.cli.discovery.RecipeCatalog;
import dev.erst.gridgrind.cli.discovery.RecipeCatalogEntry;
import dev.erst.gridgrind.cli.discovery.RecipeView;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Executes the published example requests so checked-in and built-in examples stay runnable. */
class ExampleExecutionFixturesTest {
  private static final Path GENERATED_WORKBOOKS = Path.of("generated-workbooks");

  @TempDir Path tempDir;

  @Test
  void observedTableLabelsMayBeEmptyWithoutBecomingAuthoredIdentifiers() throws IOException {
    Path source = tempDir.resolve("observed-table-name.xlsx");
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      sheet.createRow(0).createCell(0).setCellValue("Item");
      sheet.createRow(1).createCell(0).setCellValue("Observed");
      var table =
          sheet.createTable(
              new org.apache.poi.ss.util.AreaReference(
                  "A1:A2", org.apache.poi.ss.SpreadsheetVersion.EXCEL2007));
      table.setName("ValidTable");
      table.setDisplayName("ValidTable");
      table.getCTTable().setName("");
      try (OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
        {"protocolVersion":"V3","source":{"type":"EXISTING","path":"observed-table-name.xlsx"},
        "persistence":{"type":"NONE"},"steps":[{"stepId":"tables","target":{"type":"TABLE_ALL"},
        "query":{"type":"GET_TABLES"}}]}
        """);
    var result =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(request, bindingsForExample(tempDir, request)));
    var tables =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult.class,
            result.inspections().getFirst());
    assertEquals("", tables.tables().getFirst().name().value());
    assertEquals(
        result,
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(result)));
  }

  @Test
  void observedDefinedNameLabelsDoNotAcquireAuthoringRestrictions() throws IOException {
    Path source = tempDir.resolve("observed-defined-name.xlsx");
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("Observed");
      var name = workbook.createName();
      name.setNameName("ValidName");
      name.setRefersToFormula("Data!$A$1");
      try (OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    try (var archive = java.nio.file.FileSystems.newFileSystem(source)) {
      Path workbookXml = archive.getPath("/xl/workbook.xml");
      String xml = Files.readString(workbookXml);
      assertTrue(xml.contains("name=\"ValidName\""));
      Files.writeString(workbookXml, xml.replace("name=\"ValidName\"", "name=\"1invalid\""));
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
        {"protocolVersion":"V3","source":{"type":"EXISTING","path":"observed-defined-name.xlsx"},
        "persistence":{"type":"NONE"},"steps":[{"stepId":"names","target":{"type":"NAMED_RANGE_ALL"},
        "query":{"type":"GET_NAMED_RANGES"}}]}
        """);
    var result =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(request, bindingsForExample(tempDir, request)));
    var names =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.WorkbookInspectionResult.NamedRangesResult.class,
            result.inspections().getFirst());
    var observed =
        assertInstanceOf(
            dev.erst.gridgrind.core.model.formula.NamedRangeReport.RangeReport.class,
            names.namedRanges().getFirst());
    assertEquals("1invalid", observed.name().value());
    assertEquals("Data!$A$1", observed.refersToFormula());
    assertEquals(
        result,
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(result)));
  }

  @Test
  void observedMalformedValidationRangesRemainVisible() throws IOException {
    Path source = tempDir.resolve("observed-range.xlsx");
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      var validation = sheet.getCTWorksheet().addNewDataValidations().addNewDataValidation();
      var node = (org.w3c.dom.Element) validation.getDomNode();
      node.setAttribute("type", "future-validation");
      node.setAttribute("sqref", "#REF!");
      try (OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
        {"protocolVersion":"V3","source":{"type":"EXISTING","path":"observed-range.xlsx"},
        "persistence":{"type":"NONE"},"steps":[{"stepId":"ranges",
        "target":{"type":"RANGE_ALL_ON_SHEET","sheetName":"Data"},"query":{"type":"GET_DATA_VALIDATIONS"}}]}
        """);
    var result =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(request, bindingsForExample(tempDir, request)));
    var validations =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.SheetInspectionResult.DataValidationsResult.class,
            result.inspections().getFirst());
    var unsupported =
        assertInstanceOf(
            dev.erst.gridgrind.core.model.validation.DataValidationEntryReport.Unsupported.class,
            validations.validations().getFirst());
    assertEquals("#REF!", unsupported.ranges().getFirst().value());
    assertEquals("UNKNOWN", unsupported.kind());
    assertEquals(
        result,
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(result)));
  }

  @Test
  void observedWorksheetLabelsDoNotAcquireAuthoringNameRestrictions() throws IOException {
    Path source = tempDir.resolve("observed-label.xlsx");
    try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      var sheet = workbook.createSheet("Valid");
      sheet.createRow(0).createCell(0).setCellValue("Item");
      sheet.createRow(1).createCell(0).setCellValue("Observed");
      var table =
          sheet.createTable(
              new org.apache.poi.ss.util.AreaReference(
                  "A1:A2", org.apache.poi.ss.SpreadsheetVersion.EXCEL2007));
      table.setName("ObservedTable");
      table.setDisplayName("ObservedTable");
      workbook.getCTWorkbook().getSheets().getSheetArray(0).setName("invalid[]");
      try (OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    WorkbookPlan request =
        GridGrindJson.readRequest(
            """
        {"protocolVersion":"V3","source":{"type":"EXISTING","path":"observed-label.xlsx"},"persistence":{"type":"NONE"},"steps":[{"stepId":"summary","target":{"type":"WORKBOOK_CURRENT"},"query":{"type":"GET_WORKBOOK_SUMMARY"}},{"stepId":"tables","target":{"type":"TABLE_ALL"},"query":{"type":"GET_TABLES"}}]}
        """);
    WorkbookResult.Success result =
        assertInstanceOf(
            WorkbookResult.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(request, bindingsForExample(tempDir, request)));
    var summaryResult =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult.class,
            result.inspections().getFirst());
    var summary =
        assertInstanceOf(
            dev.erst.gridgrind.core.result.WorkbookSummary.WithSheets.class,
            summaryResult.workbook());
    assertEquals(
        java.util.List.of("invalid[]"),
        summary.sheetNames().stream()
            .map(dev.erst.gridgrind.core.vocabulary.SheetLabel::value)
            .toList());
    assertEquals("invalid[]", summary.activeSheetName().value());
    var tables =
        assertInstanceOf(
            dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult.class,
            result.inspections().get(1));
    assertEquals("ObservedTable", tables.tables().getFirst().name().value());
    assertEquals("invalid[]", tables.tables().getFirst().sheetName().value());
    var reopened =
        GridGrindJson.readWorkbookResult(GridGrindJsonOutput.writeWorkbookResultBytes(result));
    assertEquals(result, reopened);
  }

  @Test
  void selfContainedBuiltInExamplesExecuteFromABlankArtifactWorkspace() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("blank-artifact-workspace"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    for (RecipeCatalogEntry example : selfContainedExamples()) {
      PublishedBuiltInExample published = materializedBuiltInExample(example, workspace);
      WorkbookPlan request = published.request();
      WorkbookResult.Success success =
          assertInstanceOf(
              WorkbookResult.Success.class,
              executor.execute(request, bindingsForExample(published.workspace(), request)),
              () -> "self-contained built-in example must execute successfully: " + example.id());
      assertEquals(
          request.planId(),
          success.planId(),
          () -> "success result must retain the example plan id: " + example.id());
      assertNullFreeResponse(success, example.id());
      assertPersistedWorkbookExists(request, published.workspace());
    }
  }

  @Test
  void builtInExamplesExecuteFromARepositoryRootWorkspace() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("artifact-workspace"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    for (RecipeCatalogEntry example : exampleEntries()) {
      PublishedBuiltInExample published = materializedBuiltInExample(example, workspace);
      WorkbookPlan request = published.request();
      WorkbookResult.Success success =
          assertInstanceOf(
              WorkbookResult.Success.class,
              executor.execute(request, bindingsForExample(published.workspace(), request)),
              () -> "built-in example must execute successfully: " + example.id());
      assertEquals(
          request.planId(),
          success.planId(),
          () -> "success result must retain the example plan id: " + example.id());
      assertNullFreeResponse(success, example.id());
      assertPersistedWorkbookExists(request, published.workspace());
    }
  }

  @Test
  void repositoryAssetBackedBuiltInExamplesMaterializeAndExecuteFromABlankArtifactWorkspace()
      throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("artifact-workspace-missing-assets"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    for (RecipeCatalogEntry example : repositoryAssetBackedExamples()) {
      PublishedBuiltInExample published = materializedBuiltInExample(example, workspace);
      WorkbookResult.Success success =
          assertInstanceOf(
              WorkbookResult.Success.class,
              executor.execute(
                  published.request(),
                  bindingsForExample(published.workspace(), published.request())),
              () ->
                  "repo-asset-backed built-in example must materialize and execute: "
                      + example.id());
      assertEquals(
          published.request().planId(),
          success.planId(),
          () -> "materialized built-in example must retain its plan id");
    }
  }

  @Test
  void repositoryExamplesExecuteFromTheirOwnExamplesDirectory() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("repository-workspace"));
    Path examplesDirectory = workspace.resolve("examples");
    copyExamplesDirectory(locateRepoRoot().resolve("examples"), examplesDirectory);
    Files.createDirectories(examplesDirectory.resolve("generated-workbooks"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    for (RecipeCatalogEntry example : exampleEntries()) {
      Path requestPath = examplesDirectory.resolve(example.requestFileName());
      WorkbookPlan request = GridGrindJson.readRequest(Files.readAllBytes(requestPath));
      WorkbookResult.Success success =
          assertInstanceOf(
              WorkbookResult.Success.class,
              executor.execute(request, bindingsForExample(examplesDirectory, request)),
              () ->
                  "repository example must execute successfully in-place: "
                      + example.requestFileName());
      assertEquals(
          request.planId(),
          success.planId(),
          () -> "success result must retain the repository example plan id: " + example.id());
      assertNullFreeResponse(success, example.id());
      assertPersistedWorkbookExists(request, requestPath.getParent());
    }
  }

  private static List<RecipeCatalogEntry> exampleEntries() throws IOException {
    GridGrindCli cli = new GridGrindCli();
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        cli.run(
            new String[] {"--print-recipe-catalog"},
            InputStream.nullInputStream(),
            stdout,
            OutputStream.nullOutputStream());
    assertEquals(0, exitCode, "example catalog command must succeed");
    RecipeCatalog recipeCatalog =
        GridGrindCliJson.readBytes(stdout.toByteArray(), RecipeCatalog.class);
    return recipeCatalog.recipes().stream()
        .filter(recipe -> recipe.view() == RecipeView.EXAMPLE)
        .toList();
  }

  private static List<RecipeCatalogEntry> selfContainedExamples() throws IOException {
    return exampleEntries().stream()
        .filter(example -> example.advisory() == RecipeAdvisory.SELF_CONTAINED)
        .toList();
  }

  private static List<RecipeCatalogEntry> repositoryAssetBackedExamples() throws IOException {
    return exampleEntries().stream()
        .filter(example -> example.advisory() == RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS)
        .toList();
  }

  private static PublishedBuiltInExample materializedBuiltInExample(
      RecipeCatalogEntry example, Path workspace) throws IOException {
    GridGrindCli cli = new GridGrindCli();
    Path exampleWorkspace = workspace.resolve(example.id().toLowerCase(java.util.Locale.ROOT));
    Path requestPath = exampleWorkspace.resolve(example.requestFileName());
    int exitCode =
        cli.run(
            materializationArguments(example, exampleWorkspace, requestPath),
            InputStream.nullInputStream(),
            OutputStream.nullOutputStream(),
            OutputStream.nullOutputStream());
    assertEquals(0, exitCode, () -> "built-in example command must succeed: " + example.id());
    Files.createDirectories(exampleWorkspace.resolve(GENERATED_WORKBOOKS));
    return new PublishedBuiltInExample(
        GridGrindJson.readRequest(Files.readAllBytes(requestPath)), exampleWorkspace);
  }

  private static String[] materializationArguments(
      RecipeCatalogEntry example, Path workspace, Path requestPath) {
    if (example.advisory() == RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS) {
      return new String[] {
        "--materialize-recipe", "--lookup", example.id(), "--workspace", workspace.toString()
      };
    }
    return new String[] {
      "--print-recipe", "--lookup", example.id(), "--response", requestPath.toString()
    };
  }

  private record PublishedBuiltInExample(WorkbookPlan request, Path workspace) {}

  private static void assertNullFreeResponse(WorkbookResult response, String exampleId)
      throws IOException {
    assertTrue(
        !new String(GridGrindJsonOutput.writeWorkbookResultBytes(response), StandardCharsets.UTF_8)
            .contains(": null"),
        () -> "serialized response must omit explicit null properties: " + exampleId);
  }

  private static void assertPersistedWorkbookExists(WorkbookPlan request, Path workingDirectory) {
    String persistencePath =
        ExecutionRequestPaths.persistencePath(
            request.source(), request.persistence(), workingDirectory);
    if (persistencePath == null) {
      return;
    }
    assertTrue(
        Files.exists(Path.of(persistencePath)),
        () -> "persisted workbook must exist after example execution: " + persistencePath);
  }

  private static void copyExamplesDirectory(Path source, Path target) throws IOException {
    try (var stream = Files.walk(source)) {
      for (Path path : stream.sorted(Comparator.naturalOrder()).toList()) {
        Path relativePath = source.relativize(path);
        Path targetPath = target.resolve(relativePath);
        if (Files.isDirectory(path)) {
          Files.createDirectories(targetPath);
          continue;
        }
        Files.createDirectories(targetPath.getParent());
        Files.copy(path, targetPath);
      }
    }
  }

  private static Path locateRepoRoot() {
    Path current = Path.of("").toAbsolutePath();
    while (current != null && !Files.exists(current.resolve("settings.gradle.kts"))) {
      current = current.getParent();
    }
    if (current == null) {
      throw new AssertionError("test must run inside the GridGrind repository");
    }
    return current;
  }

  private static ExecutionInputBindings bindingsForExample(Path workspace, WorkbookPlan request) {
    if (request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile
        && existingFile
            .security()
            .flatMap(dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec::passwordRef)
            .equals(java.util.Optional.of(new SecretReference("source-open-password")))) {
      return ExecutionInputBindingsFixtureSupport.bindings(
          workspace,
          request,
          Map.of(new SecretReference("source-open-password"), "GridGrind-2026"));
    }
    return ExecutionInputBindingsFixtureSupport.bindings(workspace, request);
  }
}
