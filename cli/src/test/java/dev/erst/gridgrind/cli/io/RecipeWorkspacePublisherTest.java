package dev.erst.gridgrind.cli.io;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.discovery.Recipe;
import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import java.io.IOException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies atomic recipe-workspace publication without caller-controlled asset traversal. */
class RecipeWorkspacePublisherTest {
  @TempDir Path parent;

  @Test
  void collisionRemainsPrimaryWhenPrivateStageCleanupAlsoFails() throws Exception {
    Recipe recipe = RecipeRegistry.recipeFor("BUDGET").orElseThrow();
    Path workspace = parent.resolve("occupied-during-commit");
    var collision = new FileAlreadyExistsException(workspace.toString());
    var cleanup = new IOException("synthetic private-stage cleanup failure");
    var cleaned = new java.util.concurrent.atomic.AtomicInteger();
    var publisher =
        new RecipeWorkspacePublisher(
            (source, target) -> {
              Files.createDirectory(target);
              throw collision;
            },
            stage -> {
              cleaned.incrementAndGet();
              assertTrue(Files.isRegularFile(stage.resolve(recipe.requestFileName())));
              throw cleanup;
            });
    var failure =
        assertThrows(
            RecipeWorkspacePublicationException.class,
            () -> publisher.publishWorkspace(recipe, workspace, new byte[] {1}));
    assertTrue(failure.getMessage().contains("Workspace already exists"));
    assertSame(collision, failure.getCause());
    assertEquals(1, collision.getSuppressed().length);
    assertSame(cleanup, collision.getSuppressed()[0]);
    assertEquals(1, cleaned.get());
    try (var entries = Files.list(workspace)) {
      assertTrue(entries.findAny().isEmpty());
    }
  }

  @Test
  void assetCopyFailureRemainsPrimaryWhenPrivateStageCleanupAlsoFails() {
    var cleanup = new IOException("synthetic private-stage cleanup failure");
    var publisher =
        new RecipeWorkspacePublisher(
            (source, target) -> {
              throw new AssertionError("Invalid assets must prevent publication");
            },
            stage -> {
              throw cleanup;
            });
    IOException failure =
        assertThrows(
            IOException.class,
            () ->
                publisher.publishWorkspace(
                    recipeWithPath("missing-assets/none.xlsx"),
                    parent.resolve("copy-failed"),
                    new byte[] {1}));
    assertNotSame(cleanup, failure);
    assertEquals(1, failure.getSuppressed().length);
    assertSame(cleanup, failure.getSuppressed()[0]);
    assertFalse(Files.exists(parent.resolve("copy-failed")));
  }

  @Test
  void concurrentEmptyWorkspaceSurvivesTheFinalPublicationPrimitive() throws Exception {
    Recipe recipe = RecipeRegistry.recipeFor("BUDGET").orElseThrow();
    Path workspace = parent.resolve("concurrent-workspace");
    var identity = new java.util.concurrent.atomic.AtomicReference<Object>();
    var publisher =
        new RecipeWorkspacePublisher((source, target) -> {
          Files.createDirectory(target);
          identity.set(Files.readAttributes(target, BasicFileAttributes.class).fileKey());
          RecipeWorkspacePublisher.moveAtomically(source, target);
        });
    assertThrows(
        RecipeWorkspacePublicationException.class,
        () -> publisher.publishWorkspace(recipe, workspace, new byte[] {1}));
    assertEquals(
        identity.get(), Files.readAttributes(workspace, BasicFileAttributes.class).fileKey());
    try (var files = Files.list(workspace)) {
      assertTrue(files.findAny().isEmpty());
    }
  }

  @Test
  void publishesTheRequestAndEveryAssetOnlyInsideOneNewWorkspace() throws Exception {
    Recipe recipe = RecipeRegistry.recipeFor("AUDIT_EXISTING_WORKBOOK").orElseThrow();
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
    Recipe recipe = RecipeRegistry.recipeFor("BUDGET").orElseThrow();
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
    Recipe brokenRecipe = recipeWithPath("missing-assets/none.xlsx");
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
    Recipe recipe = RecipeRegistry.recipeFor("BUDGET").orElseThrow();

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
    Recipe recipe = RecipeRegistry.recipeFor("AUDIT_EXISTING_WORKBOOK").orElseThrow();

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

  private static Recipe recipeWithPath(String requiredPath) {
    Recipe base = RecipeRegistry.recipeFor("AUDIT_EXISTING_WORKBOOK").orElseThrow();
    return new Recipe(
        base.id(),
        base.type(),
        base.requestFileName(),
        base.summary(),
        base.advisory(),
        List.of(requiredPath),
        base.intentTags(),
        base.plan());
  }
}
