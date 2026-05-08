package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Public workbook-artifact IO facade for materialized event-read and persistence flows. */
public final class WorkbookArtifactIo {
  private WorkbookArtifactIo() {}

  /** One readable plain `.xlsx` path materialized from a plain or encrypted source workbook. */
  public interface MaterializedWorkbook extends AutoCloseable {
    /** Returns the plain materialized workbook path. */
    Path workbookPath();

    @Override
    void close() throws IOException;
  }

  /** Materializes one readable plain `.xlsx` path from a plain or encrypted source workbook. */
  public static MaterializedWorkbook materializeWorkbook(
      Path workbookPath, ExcelOoxmlOpenOptions openOptions, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    return new MaterializedWorkbookHandle(
        WorkbookPackageIo.materializeWorkbook(workbookPath, openOptions, tempFileFactory));
  }

  /**
   * Persists one materialized plain workbook with the requested encryption and signing settings.
   */
  public static void persistMaterializedWorkbook(
      Path plainWorkbookPath,
      Path targetPath,
      ExcelOoxmlPackageSecuritySnapshot sourceSecurity,
      Optional<String> sourceEncryptionPassword,
      boolean sourceMutated,
      ExcelOoxmlPersistenceOptions persistenceOptions)
      throws IOException {
    WorkbookPackageIo.persistMaterializedWorkbook(
        plainWorkbookPath,
        targetPath,
        sourceSecurity,
        sourceEncryptionPassword,
        sourceMutated,
        persistenceOptions);
  }

  private record MaterializedWorkbookHandle(WorkbookPackageIo.MaterializedWorkbook delegate)
      implements MaterializedWorkbook {
    private MaterializedWorkbookHandle {
      Objects.requireNonNull(delegate, "delegate must not be null");
    }

    @Override
    public Path workbookPath() {
      return delegate.workbookPath();
    }

    @Override
    public void close() throws IOException {
      delegate.close();
    }
  }
}
