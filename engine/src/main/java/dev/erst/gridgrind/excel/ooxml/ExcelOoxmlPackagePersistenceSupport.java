package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.UnsupportedSourceEncryptionPreservationException;

/** Resolves how a saved workbook should preserve or replace source OOXML security settings. */
public final class ExcelOoxmlPackagePersistenceSupport {
  private ExcelOoxmlPackagePersistenceSupport() {}

  /** Computes the effective encryption and signing settings for one persistence request. */
  public static ExcelOoxmlPersistenceOptions effectiveOptions(
      ExcelOoxmlPackageSecuritySnapshot sourceSecurity,
      java.util.Optional<String> sourceEncryptionPassword,
      ExcelOoxmlPersistenceOptions persistenceOptions) {
    return switch (persistenceOptions.encryption()) {
      case ExcelOoxmlPersistenceEncryption.Plaintext _ -> persistenceOptions;
      case ExcelOoxmlPersistenceEncryption.Encrypt _ -> persistenceOptions;
      case ExcelOoxmlPersistenceEncryption.PreserveSource _ ->
          new ExcelOoxmlPersistenceOptions(
              new ExcelOoxmlPersistenceEncryption.Encrypt(
                  preservedSourceEncryption(sourceSecurity, sourceEncryptionPassword)),
              persistenceOptions.signature());
    };
  }

  private static ExcelOoxmlEncryptionOptions preservedSourceEncryption(
      ExcelOoxmlPackageSecuritySnapshot sourceSecurity,
      java.util.Optional<String> sourceEncryptionPassword) { // LIM-038
    if (!(sourceSecurity.encryption()
        instanceof ExcelOoxmlEncryptionSnapshot.Encrypted sourceEncryption)) {
      throw new IllegalArgumentException(
          "PRESERVE_SOURCE encryption requires an encrypted source workbook");
    }
    if (sourceEncryptionPassword.isEmpty()) {
      throw new IllegalStateException(
          "Encrypted source workbooks must retain their verified source password while open");
    }
    if (sourceEncryption.mode()
        != dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode.AGILE) {
      throw new UnsupportedSourceEncryptionPreservationException(
          "Source workbook encryption mode "
              + sourceEncryption.mode()
              + " is readable but not auto-preservable on write;"
              + " set persistence.security.encryption explicitly to author a supported AGILE"
              + " package instead.");
    }
    if (sourceEncryption.chainingMode()
        != dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode.CBC) {
      throw new UnsupportedSourceEncryptionPreservationException(
          "Source workbook encryption chaining mode "
              + sourceEncryption.chainingMode()
              + " is readable but not auto-preservable on write;"
              + " set persistence.security.encryption explicitly to author a supported AGILE"
              + " package instead.");
    }
    return new ExcelOoxmlEncryptionOptions(
        sourceEncryptionPassword.orElseThrow(),
        ExcelOoxmlSecurityPoiBridge.toWriteCipher(sourceEncryption.cipherAlgorithm())
            .orElseThrow(
                () ->
                    new UnsupportedSourceEncryptionPreservationException(
                        "Source workbook encryption cipher "
                            + sourceEncryption.cipherAlgorithm()
                            + " is readable but not auto-preservable on write;"
                            + " set persistence.security.encryption explicitly to author a"
                            + " supported AGILE package instead.")),
        ExcelOoxmlSecurityPoiBridge.toWriteHash(sourceEncryption.hashAlgorithm())
            .orElseThrow(
                () ->
                    new UnsupportedSourceEncryptionPreservationException(
                        "Source workbook encryption hash "
                            + sourceEncryption.hashAlgorithm()
                            + " is readable but not auto-preservable on write;"
                            + " set persistence.security.encryption explicitly to author a"
                            + " supported AGILE package instead.")));
  }
}
