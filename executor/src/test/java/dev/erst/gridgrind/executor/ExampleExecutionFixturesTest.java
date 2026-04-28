package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Executes the published example requests so checked-in and built-in examples stay runnable. */
class ExampleExecutionFixturesTest {
  @TempDir Path tempDir;

  @Test
  void selfContainedBuiltInExamplesExecuteFromABlankArtifactWorkspace() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("blank-artifact-workspace"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionInputBindings workspaceBindings = new ExecutionInputBindings(workspace);
    for (GridGrindShippedExamples.ShippedExample example :
        GridGrindShippedExamples.selfContainedExamples()) {
      WorkbookPlan request = example.plan();
      GridGrindResponse.Success success =
          assertInstanceOf(
              GridGrindResponse.Success.class,
              executor.execute(request, workspaceBindings),
              () -> "self-contained built-in example must execute successfully: " + example.id());
      assertEquals(
          request.planId(),
          success.journal().planId(),
          () -> "success journal must retain the example plan id: " + example.id());
      assertNullFreeResponse(success, example.id());
      assertPersistedWorkbookExists(request, workspace);
    }
  }

  @Test
  void builtInExamplesExecuteFromARepositoryRootWorkspace() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("artifact-workspace"));
    copyExamplesDirectory(locateRepoRoot().resolve("examples"), workspace.resolve("examples"));
    Files.createDirectories(workspace.resolve("cli/build/generated-workbooks"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionInputBindings workspaceBindings = new ExecutionInputBindings(workspace);
    for (GridGrindShippedExamples.ShippedExample example : GridGrindShippedExamples.examples()) {
      WorkbookPlan request = example.plan();
      GridGrindResponse.Success success =
          assertInstanceOf(
              GridGrindResponse.Success.class,
              executor.execute(request, workspaceBindings),
              () -> "built-in example must execute successfully: " + example.id());
      assertEquals(
          request.planId(),
          success.journal().planId(),
          () -> "success journal must retain the example plan id: " + example.id());
      assertNullFreeResponse(success, example.id());
      assertPersistedWorkbookExists(request, workspace);
    }
  }

  @Test
  void repositoryAssetBackedBuiltInExamplesFailFromABlankArtifactWorkspace() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("artifact-workspace-missing-assets"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionInputBindings workspaceBindings = new ExecutionInputBindings(workspace);
    for (GridGrindShippedExamples.ShippedExample example :
        GridGrindShippedExamples.repositoryAssetBackedExamples()) {
      GridGrindResponse.Failure failure =
          assertInstanceOf(
              GridGrindResponse.Failure.class,
              executor.execute(example.plan(), workspaceBindings),
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
    Files.createDirectories(workspace.resolve("cli/build/generated-workbooks"));

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    ExecutionInputBindings exampleBindings = new ExecutionInputBindings(examplesDirectory);
    for (GridGrindShippedExamples.ShippedExample example :
        GridGrindShippedExamples.repositoryExamples()) {
      Path requestPath = examplesDirectory.resolve(example.fileName());
      WorkbookPlan request = GridGrindJson.readRequest(Files.readAllBytes(requestPath));
      GridGrindResponse.Success success =
          assertInstanceOf(
              GridGrindResponse.Success.class,
              executor.execute(request, exampleBindings),
              () -> "repository example must execute successfully in-place: " + example.fileName());
      assertEquals(
          request.planId(),
          success.journal().planId(),
          () -> "success journal must retain the repository example plan id: " + example.id());
      assertNullFreeResponse(success, example.id());
      assertPersistedWorkbookExists(request, requestPath.getParent());
    }
  }

  private static void assertNullFreeResponse(GridGrindResponse response, String exampleId)
      throws IOException {
    assertTrue(
        !new String(GridGrindJson.writeResponseBytes(response), StandardCharsets.UTF_8)
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
