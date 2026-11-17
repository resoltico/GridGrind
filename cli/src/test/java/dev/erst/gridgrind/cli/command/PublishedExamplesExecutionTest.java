package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.discovery.RecipeAdvisory;
import dev.erst.gridgrind.core.discovery.RecipeCatalogEntry;
import dev.erst.gridgrind.core.discovery.RecipeCatalogs;
import dev.erst.gridgrind.core.discovery.RecipeType;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.source.SecretReference;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
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

/** Executes published examples through the public engine API after CLI materialization. */
class PublishedExamplesExecutionTest {
  private static final Path GENERATED_WORKBOOKS = Path.of("generated-workbooks");
  @TempDir Path tempDir;

  @Test
  void selfContainedBuiltInExamplesExecuteFromABlankArtifactWorkspace() throws IOException {
    Path workspace = Files.createDirectories(tempDir.resolve("blank-artifact-workspace"));

    var executor = dev.erst.gridgrind.engine.api.GridGrindEngine.requestExecutor();
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

    var executor = dev.erst.gridgrind.engine.api.GridGrindEngine.requestExecutor();
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

    var executor = dev.erst.gridgrind.engine.api.GridGrindEngine.requestExecutor();
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

    var executor = dev.erst.gridgrind.engine.api.GridGrindEngine.requestExecutor();
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

  private static List<RecipeCatalogEntry> exampleEntries() {
    return RecipeCatalogs.catalog().recipes().stream()
        .filter(recipe -> recipe.type() == RecipeType.EXAMPLE)
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
        "recipe", "materialize", example.id(), "--workspace", workspace.toString()
      };
    }
    return new String[] {"recipe", "show", example.id(), "--response", requestPath.toString()};
  }

  private record PublishedBuiltInExample(WorkbookPlan request, Path workspace) {}

  private static void assertNullFreeResponse(WorkbookResult response, String exampleId)
      throws IOException {
    assertFalse(
        new String(GridGrindJsonOutput.writeWorkbookResultBytes(response), StandardCharsets.UTF_8)
            .contains(": null"),
        () -> "serialized response must omit explicit null properties: " + exampleId);
  }

  private static void assertPersistedWorkbookExists(WorkbookPlan request, Path workingDirectory) {
    Path path =
        switch (request.persistence()) {
          case WorkbookPlan.WorkbookPersistence.None _ -> null;
          case WorkbookPlan.WorkbookPersistence.SaveAs save ->
              workingDirectory.resolve(save.path());
          case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
              request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile source
                  ? workingDirectory.resolve(source.path())
                  : null;
        };
    if (path != null) {
      assertTrue(Files.isRegularFile(path), () -> "Persisted workbook must exist: " + path);
    }
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

  private static dev.erst.gridgrind.engine.api.GridGrindRequestInputs bindingsForExample(
      Path workspace, WorkbookPlan request) throws IOException {
    List<dev.erst.gridgrind.core.grant.ExecutionGrant.ReadAuthority> reads;
    try (var files = Files.walk(workspace)) {
      reads =
          files
              .filter(Files::isRegularFile)
              .<dev.erst.gridgrind.core.grant.ExecutionGrant.ReadAuthority>map(
                  dev.erst.gridgrind.core.grant.ExecutionGrant.ReadAuthority.File::new)
              .toList();
    }
    var publication =
        switch (request.persistence()) {
          case WorkbookPlan.WorkbookPersistence.None _ ->
              new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.None();
          case WorkbookPlan.WorkbookPersistence.SaveAs save ->
              new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.SaveAs(
                  workspace.resolve(save.path()), save.ifExists());
          case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
              new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority
                  .OverwriteSource();
        };
    var secrets =
        request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existing
            ? existing
                .security()
                .flatMap(dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec::passwordRef)
                .stream()
                .toList()
            : List.<SecretReference>of();
    var grant =
        new dev.erst.gridgrind.core.grant.ExecutionGrant(
            16L * 1024 * 1024,
            reads,
            request.steps().stream()
                .map(step ->
                    switch (step) {
                      case dev.erst.gridgrind.core.plan.MutationStep mutation ->
                          mutation.action().actionType();
                      case dev.erst.gridgrind.core.plan.InspectionStep inspection ->
                          inspection.query().queryType();
                      case dev.erst.gridgrind.core.plan.AssertionStep assertion ->
                          assertion.assertion().assertionType();
                    })
                .distinct()
                .toList(),
            new dev.erst.gridgrind.core.grant.ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            publication,
            secrets,
            dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum());
    return new dev.erst.gridgrind.engine.api.GridGrindRequestInputs(
        workspace,
        workspace.resolve(".gridgrind/tmp"),
        grant,
        reference -> {
          if (!reference.equals(new SecretReference("source-open-password"))) {
            throw new IllegalArgumentException("Unknown example secret: " + reference.id());
          }
          return "GridGrind-2026".toCharArray();
        });
  }
}
