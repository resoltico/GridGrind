package dev.erst.gridgrind.cli.examples;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies atomic recipe-workspace publication without caller-controlled asset traversal. */
class RecipeWorkspacePublisherTest {
  @TempDir Path parent;

  @Test
  void publishesTheRequestAndEveryAssetOnlyInsideOneNewWorkspace() throws Exception {
    GridGrindCliRecipe recipe =
        GridGrindCliRecipeRegistry.recipeFor("AUDIT_EXISTING_WORKBOOK").orElseThrow();
    Path workspace = parent.resolve("audit-workspace");
    byte[] requestBytes = "{\"request\":true}".getBytes(java.nio.charset.StandardCharsets.UTF_8);

    Path published = RecipeWorkspacePublisher.publish(recipe, workspace, requestBytes);

    assertTrue(Files.isSameFile(workspace, published));
    assertArrayEquals(
        requestBytes, Files.readAllBytes(workspace.resolve(recipe.requestFileName())));
    assertTrue(
        Files.isRegularFile(workspace.resolve("task-starter-assets/workbook-ops-source.xlsx")));
  }

  @Test
  void rejectsAnExistingWorkspaceSymlinkWithoutWritingThroughIt() throws Exception {
    GridGrindCliRecipe recipe = GridGrindCliRecipeRegistry.recipeFor("BUDGET").orElseThrow();
    Path outside = Files.createDirectory(parent.resolve("outside"));
    Path workspace = parent.resolve("audit-workspace");
    Files.createSymbolicLink(workspace, outside);

    assertThrows(
        RecipeWorkspacePublicationException.class,
        () -> RecipeWorkspacePublisher.publish(recipe, workspace, new byte[] {1}));
    try (var files = Files.list(outside)) {
      assertTrue(files.findAny().isEmpty());
    }
  }

  @Test
  void removesThePrivateStageWhenOnePackagedAssetCannotBeLoaded() throws Exception {
    GridGrindCliRecipe brokenRecipe = recipeWithPath("missing-assets/none.xlsx");
    Path workspace = parent.resolve("broken-workspace");

    assertThrows(
        IOException.class,
        () -> RecipeWorkspacePublisher.publish(brokenRecipe, workspace, new byte[] {1}));

    assertFalse(Files.exists(workspace));
    try (var siblings = Files.list(parent)) {
      assertTrue(
          siblings.noneMatch(
              path -> path.getFileName().toString().startsWith(".gridgrind-recipe-stage-")));
    }
  }

  @Test
  void rejectsEscapingAssetDeclarationsBeforePublishingTheWorkspace() {
    Path workspace = parent.resolve("escaping-workspace");

    assertThrows(
        IOException.class,
        () ->
            RecipeWorkspacePublisher.publish(
                recipeWithPath("../escape.xlsx"), workspace, new byte[] {1}));

    assertFalse(Files.exists(workspace));
  }

  @Test
  void translatesPublicationMoveFailuresAndCleansTheirPrivateStage() throws Exception {
    GridGrindCliRecipe recipe = GridGrindCliRecipeRegistry.recipeFor("BUDGET").orElseThrow();

    RecipeWorkspacePublicationException collision =
        assertThrows(
            RecipeWorkspacePublicationException.class,
            () ->
                new RecipeWorkspacePublisher(
                        (source, target) -> {
                          Files.delete(source.resolve(recipe.requestFileName()));
                          Files.delete(source);
                          throw new FileAlreadyExistsException(target.toString());
                        })
                    .publishWorkspace(recipe, parent.resolve("collision"), new byte[] {1}));
    RecipeWorkspacePublicationException atomicMove =
        assertThrows(
            RecipeWorkspacePublicationException.class,
            () ->
                new RecipeWorkspacePublisher(
                        (source, target) -> {
                          throw new AtomicMoveNotSupportedException(
                              source.toString(), target.toString(), "test");
                        })
                    .publishWorkspace(recipe, parent.resolve("atomic"), new byte[] {1}));

    assertTrue(collision.getMessage().contains("Workspace already exists"));
    assertTrue(atomicMove.getMessage().contains("does not support atomic recipe publication"));
    try (var siblings = Files.list(parent)) {
      assertTrue(
          siblings.noneMatch(
              path -> path.getFileName().toString().startsWith(".gridgrind-recipe-stage-")));
    }
  }

  @Test
  void rejectsAWorkspaceRootAndAMissingWorkspaceParent() {
    GridGrindCliRecipe recipe =
        GridGrindCliRecipeRegistry.recipeFor("AUDIT_EXISTING_WORKBOOK").orElseThrow();

    RecipeWorkspacePublicationException root =
        assertThrows(
            RecipeWorkspacePublicationException.class,
            () -> RecipeWorkspacePublisher.publish(recipe, Path.of("/"), new byte[] {1}));
    RecipeWorkspacePublicationException missingParent =
        assertThrows(
            RecipeWorkspacePublicationException.class,
            () ->
                RecipeWorkspacePublisher.publish(
                    recipe, parent.resolve("missing-parent/workspace"), new byte[] {1}));

    assertTrue(root.getMessage().contains("must name a new directory"));
    assertTrue(missingParent.getMessage().contains("parent does not exist"));
  }

  private static GridGrindCliRecipe recipeWithPath(String requiredPath) {
    GridGrindCliRecipe base =
        GridGrindCliRecipeRegistry.recipeFor("AUDIT_EXISTING_WORKBOOK").orElseThrow();
    return new GridGrindCliRecipe(
        base.id(),
        base.view(),
        base.requestFileName(),
        base.summary(),
        base.advisory(),
        List.of(requiredPath),
        base.intentTags(),
        base.plan());
  }
}
