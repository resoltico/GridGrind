package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.InvalidWorkbookPasswordException;
import dev.erst.gridgrind.excel.WorkbookNotOpenableException;
import dev.erst.gridgrind.excel.WorkbookPasswordRequiredException;
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
      Path workbookPath, ExcelOoxmlOpenOptions openOptions) throws IOException {
    try (ExcelOoxmlPrivateTempWorkbook privateWorkbook =
            ExcelOoxmlPrivateTempWorkbook.create("gridgrind-ooxml-decrypted-", ".xlsx");
        POIFSFileSystem fileSystem = new POIFSFileSystem(workbookPath.toFile())) {
      if (!isEncryptedOoxmlPackage(fileSystem)) {
        throw new WorkbookNotOpenableException(
            workbookPath,
            new IllegalArgumentException("Only encrypted OOXML .xlsx packages are supported"));
      }
      String password = openPassword(openOptions);
      EncryptionInfo encryptionInfo =
          ExcelOoxmlPackageEncryptionSupport.readEncryptionInfo(fileSystem, workbookPath);

      Decryptor decryptor = Decryptor.getInstance(encryptionInfo);
      boolean unlocked =
          ExcelOoxmlPackageEncryptionSupport.verifyPassword(
              decryptor::verifyPassword, password, workbookPath);
      if (!unlocked) {
        throw new InvalidWorkbookPasswordException();
      }

      ExcelOoxmlPackageEncryptionSupport.materializeDecryptedWorkbook(
          () -> decryptor.getDataStream(fileSystem), privateWorkbook.workbookPath(), workbookPath);

      ExcelOoxmlEncryptionSnapshot encryption =
          ExcelOoxmlPackageEncryptionSupport.encryptionSnapshot(encryptionInfo);
      return new ExcelOoxmlPackageSecuritySupport.ReadableWorkbook(
          privateWorkbook.workbookPath(),
          ExcelOoxmlPackageInspectionSupport.inspectPackageSecurity(
              privateWorkbook.workbookPath(), encryption),
          Optional.of(password),
          Optional.of(privateWorkbook.releaseRoot()));
    }
  }

  static void createTargetParentDirectories(Path targetPath) throws IOException {
    Path parentOrRoot = Objects.requireNonNullElse(targetPath.getParent(), targetPath.getRoot());
    Files.createDirectories(parentOrRoot);
  }

  static FileMagic fileMagic(Path workbookPath) throws IOException {
    try (java.io.InputStream inputStream = Files.newInputStream(workbookPath)) {
      return FileMagic.valueOf(FileMagic.prepareToCheckMagic(inputStream));
    }
  }

  private static boolean isEncryptedOoxmlPackage(POIFSFileSystem fileSystem) {
    return fileSystem.getRoot().hasEntryCaseInsensitive(Decryptor.DEFAULT_POIFS_ENTRY);
  }

  private static String openPassword(ExcelOoxmlOpenOptions openOptions) {
    return switch (normalizeOpenOptions(openOptions)) {
      case ExcelOoxmlOpenOptions.Unencrypted _ -> throw new WorkbookPasswordRequiredException();
      case ExcelOoxmlOpenOptions.Encrypted encrypted -> encrypted.password();
    };
  }
}
