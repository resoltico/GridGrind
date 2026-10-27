package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSecurity;
import dev.erst.gridgrind.engine.layout.LayoutAdapter;
import dev.erst.gridgrind.engine.packaging.WorkbookFiles;
import dev.erst.gridgrind.engine.packaging.WorkbookNormalization;
import dev.erst.gridgrind.engine.security.SecurityAdapter;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Sheet;

/** Workbook persistence and source-metadata operations. */
public final class WorkbookPersistence {
  private final XssfWorkbook workbook;

  WorkbookPersistence(XssfWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Saves the workbook to disk with explicit write ownership and temp-file ownership. */
  public void save(
      Path workbookPath,
      WorkbookArtifactWriteDisposition writeDisposition,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    save(workbookPath, writeDisposition, ResolvedPersistenceSecurity.none(), tempFileFactory);
  }

  /** Returns the loaded workbook path when the workbook was opened from disk. */
  public Optional<Path> sourcePath() {
    return workbook.context().sourcePath();
  }

  /** Returns the loaded OOXML package-security facts captured when the workbook was opened. */
  public OoxmlPackageSecurityReport loadedPackageSecurity() {
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

  /** Saves with explicit write disposition, package security and temporary-file ownership. */
  public void save(
      Path workbookPath,
      WorkbookArtifactWriteDisposition writeDisposition,
      ResolvedPersistenceSecurity persistenceOptions,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    Objects.requireNonNull(persistenceOptions, "persistenceOptions must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    if (persistenceOptions.writesPlaintextUnsigned()
        && !workbook.context().loadedPackageSecurity().isSecure()) {
      savePlainWorkbook(workbookPath, writeDisposition);
      return;
    }
    SecurityAdapter.saveWorkbook(
        workbook, workbookPath, writeDisposition, persistenceOptions, tempFileFactory);
  }

  /** Writes a normalized plain OOXML package under the specified write disposition. */
  public void savePlainWorkbook(
      Path workbookPath, WorkbookArtifactWriteDisposition writeDisposition) throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    Files.createDirectories(absolutePath.getParent());
    for (Sheet sheet : workbook.context().workbook()) {
      LayoutAdapter.canonicalizeColumnDefinitions((org.apache.poi.xssf.usermodel.XSSFSheet) sheet);
    }
    TableAdapter.syncAllHeaders(workbook.context().workbook());
    WorkbookNormalization.normalizeDocumentProperties(workbook.context().workbook());

    try (OutputStream outputStream =
        WorkbookFiles.newWorkbookOutputStream(absolutePath, writeDisposition)) {
      workbook.context().workbook().write(outputStream);
    }
    WorkbookNormalization.normalizeWorkbookPackage(absolutePath);
  }

  OoxmlPackageSecurityReport packageSecurity() {
    var loaded = workbook.context().loadedPackageSecurity();
    if (!workbook.persistence().wasMutatedSinceOpen() || loaded.signatures().isEmpty())
      return loaded;
    return new OoxmlPackageSecurityReport(
        loaded.encryption(),
        loaded.signatures().stream()
            .map(
                signature ->
                    new dev.erst.gridgrind.core.model.security.OoxmlSignatureReport(
                        signature.packagePartName(),
                        signature.signer(),
                        signature.state().afterMutation()))
            .toList());
  }

  void close() throws IOException {
    IOException failure = null;
    try {
      workbook.context().formulaRuntime().close();
    } catch (IOException exception) {
      failure = exception;
    }
    try {
      workbook.context().workbook().close();
    } catch (IOException exception) {
      if (failure == null) {
        failure = exception;
      } else {
        failure.addSuppressed(exception);
      }
    }
    if (failure != null) {
      throw failure;
    }
  }
}
