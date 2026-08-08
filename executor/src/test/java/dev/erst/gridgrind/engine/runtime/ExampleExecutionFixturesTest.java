package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.GridGrindCli;
import dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.discovery.RecipeCatalog;
import dev.erst.gridgrind.cli.discovery.RecipeCatalogEntry;
import dev.erst.gridgrind.cli.discovery.RecipeView;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Executes the published example requests so checked-in and built-in examples stay runnable. */
class ExampleExecutionFixturesTest {
  @TempDir Path tempDir;

  @Test
  void selfContainedBuiltInExamplesExecuteFromABlankArtifactWorkspace() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("blank-artifact-workspace"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionInputBindings workspaceBindings =
        ExecutionInputBindingsFixtureSupport.bindings(workspace);
    for (RecipeCatalogEntry example : selfContainedExamples()) {
      WorkbookPlan request = printedBuiltInExample(example.id());
      WorkbookResult.Success success =
          assertInstanceOf(
              WorkbookResult.Success.class,
              executor.execute(request, workspaceBindings),
              () -> "self-contained built-in example must execute successfully: " + example.id());
      assertEquals(
          request.planId(),
          success.planId(),
          () -> "success result must retain the example plan id: " + example.id());
      assertNullFreeResponse(success, example.id());
      assertPersistedWorkbookExists(request, workspace);
    }
  }

  @Test
  void builtInExamplesExecuteFromARepositoryRootWorkspace() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("artifact-workspace"));
    Path repositoryExamples = locateRepoRoot().resolve("examples");
    Files.createDirectories(workspace.resolve("generated-workbooks"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionInputBindings workspaceBindings =
        ExecutionInputBindingsFixtureSupport.bindings(workspace);
    for (RecipeCatalogEntry example : exampleEntries()) {
      copyRequiredExampleAssets(example, repositoryExamples, workspace);
      WorkbookPlan request = printedBuiltInExample(example.id());
      WorkbookResult.Success success =
          assertInstanceOf(
              WorkbookResult.Success.class,
              executor.execute(request, workspaceBindings),
              () -> "built-in example must execute successfully: " + example.id());
      assertEquals(
          request.planId(),
          success.planId(),
          () -> "success result must retain the example plan id: " + example.id());
      assertNullFreeResponse(success, example.id());
      assertPersistedWorkbookExists(request, workspace);
    }
  }

  @Test
  void repositoryAssetBackedBuiltInExamplesFailFromABlankArtifactWorkspace() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("artifact-workspace-missing-assets"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionInputBindings workspaceBindings =
        ExecutionInputBindingsFixtureSupport.bindings(workspace);
    for (RecipeCatalogEntry example : repositoryAssetBackedExamples()) {
      WorkbookResult.Failure failure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              executor.execute(printedBuiltInExample(example.id()), workspaceBindings),
              () ->
                  "repo-asset-backed built-in example must fail without copied assets: "
                      + example.id());
      assertEquals(
          expectedBlankWorkspaceFailureCode(example.id()),
          failure.problem().code(),
          () -> "blank artifact workspace must fail with the documented problem code");
    }
  }

  @Test
  void repositoryExamplesExecuteFromTheirOwnExamplesDirectory() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("repository-workspace"));
    Path examplesDirectory = workspace.resolve("examples");
    copyExamplesDirectory(locateRepoRoot().resolve("examples"), examplesDirectory);
    Files.createDirectories(examplesDirectory.resolve("generated-workbooks"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionInputBindings exampleBindings =
        ExecutionInputBindingsFixtureSupport.bindings(examplesDirectory);
    for (RecipeCatalogEntry example : exampleEntries()) {
      Path requestPath = examplesDirectory.resolve(example.requestFileName());
      WorkbookPlan request = GridGrindJson.readRequest(Files.readAllBytes(requestPath));
      WorkbookResult.Success success =
          assertInstanceOf(
              WorkbookResult.Success.class,
              executor.execute(request, exampleBindings),
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
        .filter(example -> example.workspaceMode() == ExampleWorkspaceMode.SELF_CONTAINED)
        .toList();
  }

  private static List<RecipeCatalogEntry> repositoryAssetBackedExamples() throws IOException {
    return exampleEntries().stream()
        .filter(example -> example.workspaceMode() == ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS)
        .toList();
  }

  private static WorkbookPlan printedBuiltInExample(String exampleId) throws IOException {
    GridGrindCli cli = new GridGrindCli();
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    int exitCode =
        cli.run(
            new String[] {"--print-recipe", "--lookup", exampleId},
            InputStream.nullInputStream(),
            stdout,
            OutputStream.nullOutputStream());
    assertEquals(0, exitCode, () -> "built-in example command must succeed: " + exampleId);
    return GridGrindJson.readRequest(stdout.toByteArray());
  }

  private static void assertNullFreeResponse(WorkbookResult response, String exampleId)
      throws IOException {
    assertTrue(
        !new String(GridGrindJsonOutput.writeWorkbookResultBytes(response), StandardCharsets.UTF_8)
            .contains(": null"),
        () -> "serialized response must omit explicit null properties: " + exampleId);
  }

  private static GridGrindProblemCode expectedBlankWorkspaceFailureCode(String exampleId) {
    return switch (exampleId) {
      case "CUSTOM_XML", "SOURCE_BACKED_INPUT" -> GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND;
      case "PACKAGE_SECURITY_INSPECTION" -> GridGrindProblemCode.WORKBOOK_NOT_FOUND;
      default ->
          throw new AssertionError("Unexpected repository-asset-backed example id: " + exampleId);
    };
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

  private static void copyRequiredExampleAssets(
      RecipeCatalogEntry example, Path repositoryExamples, Path workspace) throws IOException {
    for (String requiredPath : example.requiredWorkspacePaths()) {
      Path sourcePath = repositoryExamples.resolve(requiredPath);
      Path targetPath = workspace.resolve(requiredPath);
      Files.createDirectories(targetPath.getParent());
      Files.copy(sourcePath, targetPath);
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
}
