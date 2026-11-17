package dev.erst.gridgrind.engine.packaging;

import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** File-level copy and cleanup helpers shared by OOXML package security flows. */
public final class WorkbookFiles {
  private WorkbookFiles() {}

  /** Opens one workbook output stream under the requested write disposition. */
  public static OutputStream newWorkbookOutputStream(
      Path targetPath, WorkbookArtifactWriteDisposition writeDisposition) throws IOException {
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    return switch (writeDisposition) {
      case CREATE_NEW ->
          Files.newOutputStream(
              targetPath,
              java.nio.file.StandardOpenOption.CREATE_NEW,
              java.nio.file.StandardOpenOption.WRITE);
      case REPLACE_EXISTING ->
          Files.newOutputStream(
              targetPath,
              java.nio.file.StandardOpenOption.CREATE,
              java.nio.file.StandardOpenOption.TRUNCATE_EXISTING,
              java.nio.file.StandardOpenOption.WRITE);
    };
  }

  /** Copies one source workbook byte-for-byte to the requested target path. */
  public static void copySourceWorkbook(
      Path sourcePath, Path targetPath, WorkbookArtifactWriteDisposition writeDisposition)
      throws IOException {
    Objects.requireNonNull(sourcePath, "sourcePath must not be null");
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    if (sourcePath.equals(targetPath)
        && writeDisposition == WorkbookArtifactWriteDisposition.REPLACE_EXISTING) {
      return;
    }
    if (writeDisposition == WorkbookArtifactWriteDisposition.CREATE_NEW) {
      Files.copy(sourcePath, targetPath);
      return;
    }
    Files.copy(sourcePath, targetPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
  }

  /** Moves one workbook artifact to the requested target path under explicit write ownership. */
  public static void moveWorkbook(
      Path sourcePath, Path targetPath, WorkbookArtifactWriteDisposition writeDisposition)
      throws IOException {
    Objects.requireNonNull(sourcePath, "sourcePath must not be null");
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    if (sourcePath.equals(targetPath)) {
      return;
    }
    if (writeDisposition == WorkbookArtifactWriteDisposition.CREATE_NEW) {
      Files.move(sourcePath, targetPath);
      return;
    }
    Files.move(sourcePath, targetPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
  }

  /** Deletes one temporary path if it exists, suppressing cleanup-only failures. */
  public static void deleteIfExists(Path path) {
    if (path == null) {
      return;
    }
    try {
      Files.deleteIfExists(path);
    } catch (IOException ignored) {
      // Best-effort cleanup for executor-owned temporary files only.
    }
  }

  /** Deletes an owned tree without following links, retaining bounded cleanup failures. */
  public static void deleteTreeIfExists(Path root) throws IOException {
    deleteTreeIfExists(root, Files::deleteIfExists);
  }

  static void deleteTreeIfExists(Path root, PathDeletion deletion) throws IOException {
    if (root == null) return;
    var visitor =
        new java.nio.file.SimpleFileVisitor<Path>() {
          private @org.jspecify.annotations.Nullable IOException failure;

          private void record(IOException exception) {
            if (exception instanceof java.nio.file.NoSuchFileException) return;
            if (failure == null) failure = exception;
            else if (!failure.equals(exception) && failure.getSuppressed().length < 8)
              failure.addSuppressed(exception);
          }

          private void delete(Path path) {
            try {
              deletion.delete(path);
            } catch (IOException exception) {
              record(exception);
            }
          }

          @Override
          public java.nio.file.FileVisitResult visitFile(
              Path file, java.nio.file.attribute.BasicFileAttributes attributes) {
            delete(file);
            return java.nio.file.FileVisitResult.CONTINUE;
          }

          @Override
          public java.nio.file.FileVisitResult visitFileFailed(Path file, IOException exception) {
            record(exception);
            return java.nio.file.FileVisitResult.CONTINUE;
          }

          @Override
          public java.nio.file.FileVisitResult postVisitDirectory(
              Path directory, @org.jspecify.annotations.Nullable IOException exception) {
            if (exception != null) record(exception);
            delete(directory);
            return java.nio.file.FileVisitResult.CONTINUE;
          }
        };
    try {
      Files.walkFileTree(root, visitor);
    } catch (IOException exception) {
      visitor.record(exception);
    }
    if (visitor.failure != null) throw visitor.failure;
  }

  /** Performs one deletion while the owning traversal collects filesystem failures. */
  @FunctionalInterface
  interface PathDeletion {
    /** Deletes the supplied owned entry or reports its checked filesystem failure. */
    void delete(Path path) throws IOException;
  }
}
