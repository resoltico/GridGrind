package dev.erst.gridgrind.excel.ooxml;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** One private OS-temp workbook path owned by the OOXML security flow. */
final class ExcelOoxmlPrivateTempWorkbook implements AutoCloseable {
  private static final String ROOT_PREFIX = "gridgrind-ooxml-private-";

  private final Path root;
  private final Path workbookPath;
  private boolean released;

  private ExcelOoxmlPrivateTempWorkbook(Path root, Path workbookPath) {
    this.root = Objects.requireNonNull(root, "root must not be null");
    this.workbookPath = Objects.requireNonNull(workbookPath, "workbookPath must not be null");
  }

  static ExcelOoxmlPrivateTempWorkbook create(String prefix, String suffix) throws IOException {
    Objects.requireNonNull(prefix, "prefix must not be null");
    Objects.requireNonNull(suffix, "suffix must not be null");
    Path root = Files.createTempDirectory(ROOT_PREFIX).toAbsolutePath().normalize();
    Path workbookPath = Files.createTempFile(root, prefix, suffix).toAbsolutePath().normalize();
    return new ExcelOoxmlPrivateTempWorkbook(root, workbookPath);
  }

  Path workbookPath() {
    return workbookPath;
  }

  Path root() {
    return root;
  }

  /** Transfers cleanup ownership to the caller so close() no longer deletes the tree. */
  Path releaseRoot() {
    released = true;
    return root;
  }

  @Override
  public void close() {
    if (!released) {
      ExcelOoxmlPackageFileSupport.deleteTreeIfExists(root);
    }
  }
}
