package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.InvalidWorkbookPasswordException;
import dev.erst.gridgrind.excel.WorkbookPasswordRequiredException;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/** Shared internal helpers for OOXML package open/save security flows. */
final class ExcelOoxmlPackageSecurityInternals {
  private ExcelOoxmlPackageSecurityInternals() {}

  static ExcelOoxmlOpenOptions normalizeOpenOptions(ExcelOoxmlOpenOptions openOptions) {
    return openOptions == null ? new ExcelOoxmlOpenOptions.Unencrypted() : openOptions;
  }

  static ExcelOoxmlPackageSecuritySupport.ReadableWorkbook decryptWorkbook(
      Path workbookPath, ExcelOoxmlOpenOptions openOptions, WorkbookTempFileFactory tempFileFactory)
      throws IOException {
    Path plainWorkbookPath = tempFileFactory.createTempFile("gridgrind-ooxml-decrypted-", ".xlsx");
    boolean success = false;
    try (POIFSFileSystem fileSystem = new POIFSFileSystem(workbookPath.toFile())) {
      if (!isEncryptedOoxmlPackage(fileSystem)) {
        throw new IllegalArgumentException("Only .xlsx workbooks are supported");
      }
      String password = openPassword(workbookPath, openOptions);
      EncryptionInfo encryptionInfo =
          ExcelOoxmlPackageEncryptionSupport.readEncryptionInfo(fileSystem, workbookPath);

      Decryptor decryptor = Decryptor.getInstance(encryptionInfo);
      boolean unlocked =
          ExcelOoxmlPackageEncryptionSupport.verifyPassword(
              decryptor::verifyPassword, password, workbookPath);
      if (!unlocked) {
        throw new InvalidWorkbookPasswordException(workbookPath);
      }

      ExcelOoxmlPackageEncryptionSupport.materializeDecryptedWorkbook(
          () -> decryptor.getDataStream(fileSystem), plainWorkbookPath, workbookPath);

      ExcelOoxmlEncryptionSnapshot encryption =
          ExcelOoxmlPackageEncryptionSupport.encryptionSnapshot(encryptionInfo);
      ExcelOoxmlPackageSecuritySupport.ReadableWorkbook readableWorkbook =
          new ExcelOoxmlPackageSecuritySupport.ReadableWorkbook(
              plainWorkbookPath,
              ExcelOoxmlPackageInspectionSupport.inspectPackageSecurity(
                  plainWorkbookPath, encryption),
              Optional.of(password),
              true);
      success = true;
      return readableWorkbook;
    } finally {
      if (!success) {
        ExcelOoxmlPackageFileSupport.deleteIfExists(plainWorkbookPath);
      }
    }
  }

  static void createTargetParentDirectories(Path targetPath) throws IOException {
    Path parentOrRoot = Objects.requireNonNullElse(targetPath.getParent(), targetPath.getRoot());
    Files.createDirectories(parentOrRoot);
  }

  static boolean passThroughEligible(
      ExcelWorkbook workbook, ExcelOoxmlPersistenceOptions persistenceOptions) {
    return workbook.persistence().sourcePath().isPresent()
        && !workbook.persistence().wasMutatedSinceOpen()
        && workbook.persistence().loadedPackageSecurity().isSecure()
        && persistenceOptions.isEmpty();
  }

  static boolean requiresResigning(
      ExcelWorkbook workbook, ExcelOoxmlPersistenceOptions persistenceOptions) {
    return workbook.persistence().sourcePath().isPresent()
        && !workbook.persistence().loadedPackageSecurity().signatures().isEmpty()
        && persistenceOptions.signature().isEmpty();
  }

  static FileMagic fileMagic(Path workbookPath) throws IOException {
    try (java.io.InputStream inputStream = Files.newInputStream(workbookPath)) {
      return FileMagic.valueOf(FileMagic.prepareToCheckMagic(inputStream));
    }
  }

  private static boolean isEncryptedOoxmlPackage(POIFSFileSystem fileSystem) {
    return fileSystem.getRoot().hasEntryCaseInsensitive(Decryptor.DEFAULT_POIFS_ENTRY);
  }

  private static String openPassword(Path workbookPath, ExcelOoxmlOpenOptions openOptions) {
    return switch (normalizeOpenOptions(openOptions)) {
      case ExcelOoxmlOpenOptions.Unencrypted _ ->
          throw new WorkbookPasswordRequiredException(workbookPath);
      case ExcelOoxmlOpenOptions.Encrypted encrypted -> encrypted.password();
    };
  }
}
