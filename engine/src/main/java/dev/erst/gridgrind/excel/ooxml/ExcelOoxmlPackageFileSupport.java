package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** File-level copy and cleanup helpers shared by OOXML package security flows. */
public final class ExcelOoxmlPackageFileSupport {
  private ExcelOoxmlPackageFileSupport() {}

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
}
