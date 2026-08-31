package dev.erst.gridgrind.cli.examples;

import java.io.IOException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Objects;

/** Publishes one complete recipe workspace without traversing caller-controlled child paths. */
public final class RecipeWorkspacePublisher {
  private final WorkspaceMover workspaceMover;

  RecipeWorkspacePublisher(WorkspaceMover workspaceMover) {
    this.workspaceMover = Objects.requireNonNull(workspaceMover, "workspaceMover must not be null");
  }

  /**
   * Stages the request and assets privately, then atomically publishes one new workspace directory.
   */
  public static Path publish(GridGrindCliRecipe recipe, Path workspace, byte[] requestBytes)
      throws IOException {
    return new RecipeWorkspacePublisher(RecipeWorkspacePublisher::moveAtomically)
        .publishWorkspace(recipe, workspace, requestBytes);
  }

  Path publishWorkspace(GridGrindCliRecipe recipe, Path workspace, byte[] requestBytes)
      throws IOException {
    Objects.requireNonNull(recipe, "recipe must not be null");
    Objects.requireNonNull(workspace, "workspace must not be null");
    Objects.requireNonNull(requestBytes, "requestBytes must not be null");
    Path target = normalizedTarget(workspace);
    requireAbsent(target);
    Path stage = Files.createTempDirectory(target.getParent(), ".gridgrind-recipe-stage-");
    try {
      RecipeAssetMaterializer.copyToStaging(recipe, stage);
      Files.write(stage.resolve(recipe.requestFileName()), requestBytes);
      workspaceMover.move(stage, target);
      return target;
    } catch (FileAlreadyExistsException exception) {
      throw new RecipeWorkspacePublicationException(
          "Workspace already exists: " + target, exception);
    } catch (AtomicMoveNotSupportedException exception) {
      throw new RecipeWorkspacePublicationException(
          "Workspace parent does not support atomic recipe publication: " + target.getParent(),
          exception);
    } finally {
      deleteStage(stage);
    }
  }

  private static void moveAtomically(Path source, Path target) throws IOException {
    Files.move(source, target, java.nio.file.StandardCopyOption.ATOMIC_MOVE);
  }

  private static Path normalizedTarget(Path workspace) throws IOException {
    Path authored = workspace.toAbsolutePath().normalize();
    Path fileName = authored.getFileName();
    if (fileName == null) {
      throw new RecipeWorkspacePublicationException(
          "Workspace must name a new directory beneath an existing parent: " + workspace);
    }
    Path parent =
        Objects.requireNonNull(
            authored.getParent(), "an absolute non-root workspace path must have a parent");
    try {
      return parent.toRealPath().resolve(fileName);
    } catch (java.nio.file.NoSuchFileException exception) {
      throw new RecipeWorkspacePublicationException(
          "Workspace parent does not exist: " + parent, exception);
    }
  }

  private static void requireAbsent(Path target) throws RecipeWorkspacePublicationException {
    if (Files.exists(target, LinkOption.NOFOLLOW_LINKS)) {
      throw new RecipeWorkspacePublicationException("Workspace already exists: " + target);
    }
  }

  private static void deleteStage(Path stage) throws IOException {
    if (!Files.exists(stage, LinkOption.NOFOLLOW_LINKS)) {
      return;
    }
    Files.walkFileTree(
        stage,
        new SimpleFileVisitor<>() {
          @Override
          public FileVisitResult visitFile(Path file, BasicFileAttributes attributes)
              throws IOException {
            Files.delete(file);
            return FileVisitResult.CONTINUE;
          }

          @Override
          public FileVisitResult postVisitDirectory(Path directory, IOException exception)
              throws IOException {
            Files.delete(directory);
            return FileVisitResult.CONTINUE;
          }
        });
  }

  /** Moves one staged directory to its final workspace path. */
  @FunctionalInterface
  interface WorkspaceMover {
    /** Moves one stage directory atomically when the local filesystem supports it. */
    void move(Path source, Path target) throws IOException;
  }
}
