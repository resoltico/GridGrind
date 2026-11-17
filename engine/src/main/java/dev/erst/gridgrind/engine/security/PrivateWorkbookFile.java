package dev.erst.gridgrind.engine.security;

import dev.erst.gridgrind.engine.packaging.WorkbookFiles;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** One private invocation-owned workbook path owned by the OOXML security flow. */
final class PrivateWorkbookFile implements AutoCloseable {
  private static final String ROOT_PREFIX = "gridgrind-ooxml-private-";

  private final Path root;
  private final Path workbookPath;
  private boolean released;

  private PrivateWorkbookFile(Path root, Path workbookPath) {
    this.root = Objects.requireNonNull(root, "root must not be null");
    this.workbookPath = Objects.requireNonNull(workbookPath, "workbookPath must not be null");
  }

  static PrivateWorkbookFile create(
      String prefix,
      String suffix,
      dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory factory)
      throws IOException {
    Objects.requireNonNull(prefix, "prefix must not be null");
    Objects.requireNonNull(suffix, "suffix must not be null");
    Path root = factory.createTempFile(ROOT_PREFIX, ".directory").toAbsolutePath().normalize();
    try {
      Files.delete(root);
      Files.createDirectory(root);
      Path workbookPath = Files.createTempFile(root, prefix, suffix).toAbsolutePath().normalize();
      return new PrivateWorkbookFile(root, workbookPath);
    } catch (IOException | RuntimeException | Error failure) {
      try {
        WorkbookFiles.deleteTreeIfExists(root);
      } catch (IOException cleanup) {
        if (failure.getSuppressed().length < 8) failure.addSuppressed(cleanup);
      }
      throw failure;
    }
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
  public void close() throws IOException {
    if (!released) {
      WorkbookFiles.deleteTreeIfExists(root);
    }
  }
}
