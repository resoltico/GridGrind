package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Workbook persistence and source-metadata operations. */
public final class ExcelWorkbookPersistence {
  private final ExcelWorkbook workbook;

  ExcelWorkbookPersistence(ExcelWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Saves the workbook to disk with explicit write ownership and temp-file ownership. */
  public void save(
      Path workbookPath,
      WorkbookArtifactWriteDisposition writeDisposition,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    save(workbookPath, writeDisposition, ExcelOoxmlPersistenceOptions.none(), tempFileFactory);
  }

  /** Saves the workbook to disk with explicit package-security and write ownership. */
  public void save(
      Path workbookPath,
      WorkbookArtifactWriteDisposition writeDisposition,
      ExcelOoxmlPersistenceOptions persistenceOptions,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    ExcelWorkbookPersistenceSupport.save(
        workbook, workbookPath, writeDisposition, persistenceOptions, tempFileFactory);
  }

  /** Saves the plain OOXML workbook package with explicit write ownership. */
  public void savePlainWorkbook(
      Path workbookPath, WorkbookArtifactWriteDisposition writeDisposition) throws IOException {
    ExcelWorkbookPersistenceSupport.savePlainWorkbook(workbook, workbookPath, writeDisposition);
  }

  /** Returns the loaded workbook path when the workbook was opened from disk. */
  public Optional<Path> sourcePath() {
    return workbook.context().sourcePath();
  }

  /** Returns the loaded OOXML package-security facts captured when the workbook was opened. */
  public ExcelOoxmlPackageSecuritySnapshot loadedPackageSecurity() {
    return workbook.context().loadedPackageSecurity();
  }

  /** Returns the source-open encryption password when the workbook was opened encrypted. */
  public Optional<String> sourceEncryptionPassword() {
    return workbook.context().sourceEncryptionPassword();
  }

  /** Returns whether the current in-memory workbook package differs from the loaded source. */
  public boolean wasMutatedSinceOpen() {
    return workbook.wasMutatedSinceOpenInternal();
  }
}
