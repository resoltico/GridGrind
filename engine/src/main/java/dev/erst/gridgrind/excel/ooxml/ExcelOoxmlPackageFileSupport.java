package dev.erst.gridgrind.excel.ooxml;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** File-level copy and cleanup helpers shared by OOXML package security flows. */
public final class ExcelOoxmlPackageFileSupport {
  private ExcelOoxmlPackageFileSupport() {}

  /** Copies one source workbook byte-for-byte to the requested target path. */
  public static void copySourceWorkbook(Path sourcePath, Path targetPath) throws IOException {
    Objects.requireNonNull(sourcePath, "sourcePath must not be null");
    if (sourcePath.equals(targetPath)) {
      return;
    }
    Files.copy(sourcePath, targetPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
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
