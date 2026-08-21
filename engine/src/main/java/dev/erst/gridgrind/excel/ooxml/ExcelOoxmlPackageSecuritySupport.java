package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.ExcelDeterministicWorkbookArtifactSupport;
import dev.erst.gridgrind.excel.ExcelFormulaEnvironment;
import dev.erst.gridgrind.excel.ExcelTempFileWriteTargetSupport;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.WorkbookNotOpenableException;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.poifs.filesystem.FileMagic;

/**
 * OOXML package-security bridge for encrypted open/save and package-signature inspection/signing.
 */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelOoxmlPackageSecuritySupport {
  private ExcelOoxmlPackageSecuritySupport() {}

  /** Opens one workbook through the full formula runtime plus any required package security. */
  public static ExcelWorkbook openWorkbook(
      Path workbookPath,
      ExcelFormulaEnvironment formulaEnvironment,
      ExcelOoxmlOpenOptions openOptions,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    try (ReadableWorkbook materialized =
        materializeReadableWorkbook(workbookPath, openOptions, tempFileFactory)) {
      return ExcelWorkbooks.openMaterializedWorkbook(
          materialized.workbookPath(),
          formulaEnvironment,
          Optional.of(workbookPath.toAbsolutePath().normalize()),
          materialized.packageSecurity(),
          materialized.sourceEncryptionPassword());
    }
  }

  /**
   * Opens one workbook without an explicit formula environment but with package security support.
   */
  public static ExcelWorkbook openWorkbook(
      Path workbookPath, ExcelOoxmlOpenOptions openOptions, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    try (ReadableWorkbook materialized =
        materializeReadableWorkbook(workbookPath, openOptions, tempFileFactory)) {
      return ExcelWorkbooks.openMaterializedWorkbook(
          materialized.workbookPath(),
          Optional.of(workbookPath.toAbsolutePath().normalize()),
          materialized.packageSecurity(),
          materialized.sourceEncryptionPassword());
    }
  }

  /** Materializes one readable plain `.xlsx` path from a plain or encrypted source workbook. */
  public static ReadableWorkbook materializeReadableWorkbook(
      Path workbookPath, ExcelOoxmlOpenOptions openOptions, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    ExcelOoxmlOpenOptions effectiveOpenOptions =
        ExcelOoxmlPackageSecurityInternals.normalizeOpenOptions(openOptions);

    Path absolutePath = workbookPath.toAbsolutePath().normalize();
    if (!Files.exists(absolutePath)) {
      throw new WorkbookNotFoundException(absolutePath);
    }

    FileMagic fileMagic = ExcelOoxmlPackageSecurityInternals.fileMagic(absolutePath);
    return switch (fileMagic) {
      case OOXML ->
          new ReadableWorkbook(
              absolutePath,
              ExcelOoxmlPackageInspectionSupport.inspectPackageSecurity(
                  absolutePath, ExcelOoxmlEncryptionSnapshot.none()),
              Optional.empty(),
              Optional.empty());
      case OLE2 ->
          ExcelOoxmlPackageSecurityInternals.decryptWorkbook(absolutePath, effectiveOpenOptions);
      default ->
          throw new WorkbookNotOpenableException(
              absolutePath,
              new IllegalArgumentException(
                  "Only .xlsx workbooks are supported; unsupported package magic"));
    };
  }

  public static void saveWorkbook(
      ExcelWorkbook workbook,
      Path targetPath,
      WorkbookArtifactWriteDisposition writeDisposition,
      ExcelOoxmlPersistenceOptions persistenceOptions,
      WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");

    Path normalizedTarget = targetPath.toAbsolutePath().normalize();
    ExcelOoxmlPackageSecurityInternals.createTargetParentDirectories(normalizedTarget);

    ExcelOoxmlPersistenceOptions explicitOptions =
        persistenceOptions == null ? ExcelOoxmlPersistenceOptions.none() : persistenceOptions;

    ExcelOoxmlPersistenceOptions effectiveOptions =
        ExcelOoxmlPackagePersistenceSupport.effectiveOptions(
            workbook.persistence().loadedPackageSecurity(),
            workbook.persistence().sourceEncryptionPassword(),
            explicitOptions);
    if (effectiveOptions.writesPlaintextUnsigned()) {
      workbook.persistence().savePlainWorkbook(normalizedTarget, writeDisposition);
      ExcelOoxmlPackageSigningSupport.removeSignatures(normalizedTarget);
      ExcelDeterministicWorkbookArtifactSupport.normalizeWorkbookPackage(normalizedTarget);
      return;
    }

    try (ExcelOoxmlPrivateTempWorkbook privateWorkbook =
        ExcelOoxmlPrivateTempWorkbook.create("gridgrind-ooxml-security-", ".xlsx")) {
      Path plainWorkbookPath =
          ExcelTempFileWriteTargetSupport.prepareCreateNewTarget(privateWorkbook.workbookPath());
      workbook
          .persistence()
          .savePlainWorkbook(plainWorkbookPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      persistMaterializedWorkbook(
          plainWorkbookPath,
          normalizedTarget,
          workbook.persistence().loadedPackageSecurity(),
          workbook.persistence().sourceEncryptionPassword(),
          writeDisposition,
          effectiveOptions);
    }
  }

  /**
   * Persists one materialized plain workbook with the requested encryption and signing settings.
   */
  public static void persistMaterializedWorkbook(
      Path plainWorkbookPath,
      Path targetPath,
      ExcelOoxmlPackageSecuritySnapshot sourceSecurity,
      Optional<String> sourceEncryptionPassword,
      WorkbookArtifactWriteDisposition writeDisposition,
      ExcelOoxmlPersistenceOptions persistenceOptions)
      throws IOException {
    Objects.requireNonNull(plainWorkbookPath, "plainWorkbookPath must not be null");
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(sourceSecurity, "sourceSecurity must not be null");
    Objects.requireNonNull(writeDisposition, "writeDisposition must not be null");
    Objects.requireNonNull(persistenceOptions, "persistenceOptions must not be null");

    Path normalizedTarget = targetPath.toAbsolutePath().normalize();
    ExcelOoxmlPackageSecurityInternals.createTargetParentDirectories(normalizedTarget);

    ExcelOoxmlPersistenceOptions effectiveOptions =
        ExcelOoxmlPackagePersistenceSupport.effectiveOptions(
            sourceSecurity, sourceEncryptionPassword, persistenceOptions);
    ExcelOoxmlPackageSigningSupport.removeSignatures(plainWorkbookPath);
    ExcelDeterministicWorkbookArtifactSupport.normalizeWorkbookPackage(plainWorkbookPath);
    if (effectiveOptions.signature() instanceof ExcelOoxmlPersistenceSignature.Sign sign) {
      ExcelOoxmlPackageSigningSupport.signWorkbook(plainWorkbookPath, sign.options());
    }
    if (effectiveOptions.encryption() instanceof ExcelOoxmlPersistenceEncryption.Encrypt encrypt) {
      ExcelOoxmlPackageEncryptionSupport.encryptWorkbook(
          plainWorkbookPath, normalizedTarget, writeDisposition, encrypt.options());
    } else {
      ExcelOoxmlPackageFileSupport.moveWorkbook(
          plainWorkbookPath, normalizedTarget, writeDisposition);
    }
  }

  /** One readable plain `.xlsx` path materialized from a possibly encrypted source workbook. */
  public static final class ReadableWorkbook implements AutoCloseable {
    private final Path workbookPath;
    private final ExcelOoxmlPackageSecuritySnapshot packageSecurity;
    private final Optional<String> sourceEncryptionPassword;
    private final Optional<Path> cleanupRoot;

    ReadableWorkbook(
        Path workbookPath,
        ExcelOoxmlPackageSecuritySnapshot packageSecurity,
        Optional<String> sourceEncryptionPassword,
        Optional<Path> cleanupRoot) {
      this.workbookPath = Objects.requireNonNull(workbookPath, "workbookPath must not be null");
      this.packageSecurity =
          Objects.requireNonNull(packageSecurity, "packageSecurity must not be null");
      this.sourceEncryptionPassword =
          Objects.requireNonNull(
              sourceEncryptionPassword, "sourceEncryptionPassword must not be null");
      this.cleanupRoot = Objects.requireNonNull(cleanupRoot, "cleanupRoot must not be null");
    }

    public Path workbookPath() {
      return workbookPath;
    }

    public ExcelOoxmlPackageSecuritySnapshot packageSecurity() {
      return packageSecurity;
    }

    public Optional<String> sourceEncryptionPassword() {
      return sourceEncryptionPassword;
    }

    @Override
    public void close() {
      if (cleanupRoot.isPresent()) {
        ExcelOoxmlPackageFileSupport.deleteTreeIfExists(cleanupRoot.orElseThrow());
      }
    }
  }
}
