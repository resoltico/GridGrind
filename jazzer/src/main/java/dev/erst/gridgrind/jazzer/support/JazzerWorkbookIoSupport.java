package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.runtime.ExecutionInputBindings;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** Explicit-root workbook IO helpers shared across Jazzer replay and round-trip support. */
public final class JazzerWorkbookIoSupport {
  private static final Path MANAGED_TEMP_SEGMENT = Path.of(".gridgrind", "tmp");

  private JazzerWorkbookIoSupport() {}

  /** Returns execution bindings rooted at one explicit Jazzer-owned working directory. */
  public static ExecutionInputBindings executionBindings(Path workingDirectory) {
    Path normalizedWorkingDirectory = normalize(workingDirectory, "workingDirectory");
    return new ExecutionInputBindings(
        normalizedWorkingDirectory,
        normalizedWorkingDirectory.resolve(MANAGED_TEMP_SEGMENT),
        new ExecutionGrant(
            List.of(),
            List.of(),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum()));
  }

  /** Opens one workbook using a temp factory rooted next to the workbook path. */
  public static XssfWorkbook openWorkbook(Path workbookPath) throws IOException {
    return XssfWorkbooks.open(workbookPath, tempFileFactoryFor(workbookPath));
  }

  /** Saves one workbook using a temp factory rooted next to the workbook path. */
  public static void saveWorkbook(XssfWorkbook workbook, Path workbookPath) throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    workbook
        .persistence()
        .save(
            workbookPath,
            WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
            tempFileFactoryFor(workbookPath));
  }

  /** Returns a temp-file factory rooted beside one anchored workbook path. */
  public static WorkbookTempFileFactory tempFileFactoryFor(Path anchoredPath) {
    Path normalizedAnchoredPath = normalize(anchoredPath, "anchoredPath");
    Path parent = normalizedAnchoredPath.getParent();
    Path workingDirectory = parent == null ? normalizedAnchoredPath : parent;
    return WorkbookTempFileFactory.rooted(workingDirectory.resolve(MANAGED_TEMP_SEGMENT));
  }

  private static Path normalize(Path path, String name) {
    return Objects.requireNonNull(path, name + " must not be null").toAbsolutePath().normalize();
  }
}
