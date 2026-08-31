package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies that asset-backed recipes materialize their complete runnable workspace. */
class GridGrindCliAssetRecipeMaterializationTest {
  @TempDir Path tempDir;

  @Test
  void directsAssetBackedRecipePrintingToAtomicWorkspaceMaterialization() throws Exception {
    Path responsePath = tempDir.resolve("audit-request.json");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--print-recipe",
                  "--lookup",
                  "AUDIT_EXISTING_WORKBOOK",
                  "--response",
                  responsePath.toString()
                },
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(2, exitCode);
    assertEquals(
        GridGrindProblemCode.INVALID_ARGUMENTS,
        GridGrindCliJson.readBytes(
                stdout.toByteArray(), dev.erst.gridgrind.cli.discovery.CommandError.class)
            .primaryProblem()
            .code());
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
    assertTrue(stdout.toString(StandardCharsets.UTF_8).contains("--materialize-recipe"));
    assertFalse(Files.exists(responsePath));
  }

  @Test
  void materializesPackageSecurityExampleAtomically() throws Exception {
    Path parent = Files.createTempDirectory("gridgrind-package-recipe-");
    Path workspace = parent.resolve("package-security-workspace");
    Path requestPath =
        workspace.resolve(
            GridGrindCliRecipeRegistry.recipeFor("PACKAGE_SECURITY_INSPECTION")
                .orElseThrow()
                .requestFileName());
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--materialize-recipe",
                  "--lookup",
                  "PACKAGE_SECURITY_INSPECTION",
                  "--workspace",
                  workspace.toString()
                },
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(0, exitCode);
    assertEquals(
        GridGrindShippedExamples.find("PACKAGE_SECURITY_INSPECTION").orElseThrow().plan(),
        GridGrindJson.readRequest(Files.readAllBytes(requestPath)));
    assertEquals(
        GridGrindShippedExamples.find("PACKAGE_SECURITY_INSPECTION").orElseThrow().plan(),
        GridGrindJson.readRequest(stdout.toByteArray()));
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
    assertTrue(
        Files.isRegularFile(
            workspace.resolve("package-security-assets/gridgrind-package-security.xlsx")));
  }

  @Test
  void materializesAuditTaskStarterIntoItsCataloguedRequestFile() throws Exception {
    Path parent = Files.createTempDirectory("gridgrind-task-recipe-");
    Path workspace = parent.resolve("audit-workspace");
    Path requestPath =
        workspace.resolve(
            GridGrindCliRecipeRegistry.recipeFor("AUDIT_EXISTING_WORKBOOK")
                .orElseThrow()
                .requestFileName());
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {
                  "--materialize-recipe",
                  "--lookup",
                  "AUDIT_EXISTING_WORKBOOK",
                  "--workspace",
                  workspace.toString()
                },
                InputStream.nullInputStream(),
                stdout,
                stderr);

    assertEquals(0, exitCode);
    assertEquals(
        GridGrindCliRecipeRegistry.recipeFor("AUDIT_EXISTING_WORKBOOK").orElseThrow().plan(),
        GridGrindJson.readRequest(Files.readAllBytes(requestPath)));
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
    assertTrue(
        Files.isRegularFile(workspace.resolve("task-starter-assets/workbook-ops-source.xlsx")));
  }

  @Test
  void reportsUnknownRecipesAndExistingWorkspacesAsActionableCommandErrors() throws Exception {
    Path parent = Files.createTempDirectory("gridgrind-materialize-recipe-");
    Path workspace = Files.createDirectory(parent.resolve("existing-workspace"));
    ByteArrayOutputStream unknownStdout = new ByteArrayOutputStream();
    ByteArrayOutputStream existingStdout = new ByteArrayOutputStream();

    int unknownExitCode =
        new GridGrindCli()
            .run(
                materializeArguments("UNKNOWN", parent.resolve("unknown")),
                InputStream.nullInputStream(),
                unknownStdout);
    int existingExitCode =
        new GridGrindCli()
            .run(
                materializeArguments("BUDGET", workspace),
                InputStream.nullInputStream(),
                existingStdout);

    assertEquals(2, unknownExitCode);
    assertEquals(2, existingExitCode);
    assertEquals(
        GridGrindProblemCode.INVALID_ARGUMENTS,
        GridGrindCliJson.readBytes(
                unknownStdout.toByteArray(), dev.erst.gridgrind.cli.discovery.CommandError.class)
            .primaryProblem()
            .code());
    assertTrue(
        GridGrindCliJson.readBytes(
                existingStdout.toByteArray(), dev.erst.gridgrind.cli.discovery.CommandError.class)
            .primaryProblem()
            .message()
            .contains("Workspace already exists"));
    try (var files = Files.list(workspace)) {
      assertFalse(files.findAny().isPresent());
    }
  }

  private static String[] materializeArguments(String recipeId, Path workspace) {
    return List.of(
            "--materialize-recipe", "--lookup", recipeId, "--workspace", workspace.toString())
        .toArray(String[]::new);
  }
}
