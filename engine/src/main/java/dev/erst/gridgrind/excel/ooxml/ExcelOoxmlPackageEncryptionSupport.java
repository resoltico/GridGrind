package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.WorkbookNotOpenableException;
import dev.erst.gridgrind.excel.WorkbookSecurityException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import org.apache.poi.poifs.crypt.CipherAlgorithm;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/** Encryption and decryption seams for OOXML package materialization and persistence. */
public final class ExcelOoxmlPackageEncryptionSupport {
  private ExcelOoxmlPackageEncryptionSupport() {}

  /** Reads one encrypted workbook's factual OOXML encryption metadata. */
  public static ExcelOoxmlEncryptionSnapshot encryptionSnapshot(EncryptionInfo encryptionInfo)
      throws IOException {
    try {
      return new ExcelOoxmlEncryptionSnapshot.Encrypted(
          ExcelOoxmlSecurityPoiBridge.fromPoi(encryptionInfo.getEncryptionMode()),
          ExcelOoxmlSecurityPoiBridge.fromPoi(encryptionInfo.getHeader().getCipherAlgorithm()),
          ExcelOoxmlSecurityPoiBridge.fromPoi(encryptionInfo.getVerifier().getHashAlgorithm()),
          ExcelOoxmlSecurityPoiBridge.fromPoi(encryptionInfo.getHeader().getChainingMode()),
          encryptionInfo.getHeader().getKeySize(),
          encryptionInfo.getHeader().getBlockSize(),
          encryptionInfo.getVerifier().getSpinCount());
    } catch (RuntimeException exception) {
      throw new WorkbookSecurityException("Failed to inspect OOXML encryption metadata", exception);
    }
  }

  /** Loads POI's encryption envelope from one encrypted OOXML package container. */
  public static EncryptionInfo readEncryptionInfo(POIFSFileSystem fileSystem, Path workbookPath)
      throws IOException {
    try {
      return new EncryptionInfo(fileSystem);
    } catch (IOException exception) {
      throw new WorkbookNotOpenableException(workbookPath, exception);
    }
  }

  /** Verifies one candidate workbook password through a checked-exception seam. */
  public static boolean verifyPassword(
      PasswordVerifier passwordVerifier, String password, Path workbookPath) throws IOException {
    try {
      return passwordVerifier.verifyPassword(password);
    } catch (GeneralSecurityException exception) {
      throw new WorkbookSecurityException(
          "Failed to verify the encrypted workbook password: " + workbookPath, exception);
    }
  }

  /** Materializes decrypted OOXML bytes onto one temporary plain workbook path. */
  public static void materializeDecryptedWorkbook(
      DecryptedWorkbookStreamSupplier decryptedWorkbookStreamSupplier,
      Path plainWorkbookPath,
      Path workbookPath)
      throws IOException {
    try (java.io.InputStream decryptedStream = decryptedWorkbookStreamSupplier.open();
        java.io.OutputStream outputStream = Files.newOutputStream(plainWorkbookPath)) {
      decryptedStream.transferTo(outputStream);
    } catch (GeneralSecurityException exception) {
      throw new WorkbookSecurityException(
          "Failed to decrypt the OOXML workbook package: " + workbookPath, exception);
    }
  }

  /** Encrypts one plain materialized workbook into the requested OOXML encryption mode. */
  public static void encryptWorkbook(
      Path plainWorkbookPath,
      Path targetPath,
      WorkbookArtifactWriteDisposition writeDisposition,
      ExcelOoxmlEncryptionOptions encryptionOptions)
      throws IOException {
    CipherAlgorithm cipherAlgorithm = ExcelOoxmlSecurityPoiBridge.toPoi(encryptionOptions.cipher());
    EncryptionInfo encryptionInfo =
        new EncryptionInfo(
            EncryptionMode.agile,
            cipherAlgorithm,
            ExcelOoxmlSecurityPoiBridge.toPoi(encryptionOptions.hash()),
            cipherAlgorithm.defaultKeySize,
            cipherAlgorithm.blockSize,
            ExcelOoxmlSecurityPoiBridge.toPoi(
                dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode.CBC));
    Encryptor encryptor = encryptionInfo.getEncryptor();
    encryptor.confirmPassword(encryptionOptions.password());

    writeEncryptedWorkbook(
        fileSystem -> encryptor.getDataStream(fileSystem),
        plainWorkbookPath,
        targetPath,
        writeDisposition);
  }

  /** Writes one encrypted POIFS-backed OOXML package to the target workbook path. */
  public static void writeEncryptedWorkbook(
      EncryptedWorkbookStreamSupplier encryptedWorkbookStreamSupplier,
      Path plainWorkbookPath,
      Path targetPath,
      WorkbookArtifactWriteDisposition writeDisposition)
      throws IOException {
    try (POIFSFileSystem fileSystem = new POIFSFileSystem()) {
      try (java.io.OutputStream encryptedStream =
          encryptedWorkbookStreamSupplier.open(fileSystem)) {
        Files.copy(plainWorkbookPath, encryptedStream);
      }
      try (java.io.OutputStream outputStream =
          ExcelOoxmlPackageFileSupport.newWorkbookOutputStream(targetPath, writeDisposition)) {
        fileSystem.writeFilesystem(outputStream);
      }
    } catch (GeneralSecurityException exception) {
      throw new WorkbookSecurityException(
          "Failed to encrypt the saved OOXML workbook package: " + targetPath, exception);
    }
  }

  /** Minimal password-verification seam used to isolate checked decryption failures in tests. */
  @FunctionalInterface
  public interface PasswordVerifier {
    /** Verifies one candidate workbook password. */
    boolean verifyPassword(String password) throws GeneralSecurityException;
  }

  /** Supplies one decrypted OOXML input stream for materialization. */
  @FunctionalInterface
  public interface DecryptedWorkbookStreamSupplier {
    /** Opens the decrypted workbook bytes. */
    java.io.InputStream open() throws IOException, GeneralSecurityException;
  }

  /** Supplies one encrypted OOXML output stream backed by a POIFS container. */
  @FunctionalInterface
  public interface EncryptedWorkbookStreamSupplier {
    /** Opens the encrypted output stream for one POIFS package. */
    java.io.OutputStream open(POIFSFileSystem fileSystem)
        throws IOException, GeneralSecurityException;
  }
}
