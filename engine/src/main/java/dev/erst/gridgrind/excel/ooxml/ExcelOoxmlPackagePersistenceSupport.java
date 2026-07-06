package dev.erst.gridgrind.excel.ooxml;

import java.util.Optional;

/** Resolves how a saved workbook should preserve or replace source OOXML security settings. */
public final class ExcelOoxmlPackagePersistenceSupport {
  private ExcelOoxmlPackagePersistenceSupport() {}

  /** Computes the effective encryption and signing settings for one persistence request. */
  public static ExcelOoxmlPersistenceOptions effectiveOptions(
      ExcelOoxmlPackageSecuritySnapshot sourceSecurity,
      Optional<String> sourceEncryptionPassword,
      ExcelOoxmlPersistenceOptions persistenceOptions) {
    Optional<ExcelOoxmlEncryptionOptions> encryption = persistenceOptions.encryption();
    if (encryption.isEmpty()
        && sourceSecurity.encryption()
            instanceof ExcelOoxmlEncryptionSnapshot.Encrypted sourceEncryption) {
      if (sourceEncryptionPassword.isEmpty()) {
        throw new IllegalStateException(
            "Encrypted source workbooks must retain their verified source password while open");
      }
      encryption =
          Optional.of(
              preservedSourceEncryption(sourceEncryptionPassword.orElseThrow(), sourceEncryption));
    }
    return new ExcelOoxmlPersistenceOptions(encryption, persistenceOptions.signature());
  }

  private static ExcelOoxmlEncryptionOptions preservedSourceEncryption(
      String password, ExcelOoxmlEncryptionSnapshot.Encrypted sourceEncryption) { // LIM-038
    if (sourceEncryption.mode()
        != dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode.AGILE) {
      throw new IllegalArgumentException(
          "Source workbook encryption mode "
              + sourceEncryption.mode()
              + " is readable but not auto-preservable on write;"
              + " set persistence.security.encryption explicitly to author a supported AGILE"
              + " package instead.");
    }
    if (sourceEncryption.chainingMode()
        != dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode.CBC) {
      throw new IllegalArgumentException(
          "Source workbook encryption chaining mode "
              + sourceEncryption.chainingMode()
              + " is readable but not auto-preservable on write;"
              + " set persistence.security.encryption explicitly to author a supported AGILE"
              + " package instead.");
    }
    return new ExcelOoxmlEncryptionOptions(
        password,
        ExcelOoxmlSecurityPoiBridge.toWriteCipher(sourceEncryption.cipherAlgorithm())
            .orElseThrow(
                () ->
                    new IllegalArgumentException(
                        "Source workbook encryption cipher "
                            + sourceEncryption.cipherAlgorithm()
                            + " is readable but not auto-preservable on write;"
                            + " set persistence.security.encryption explicitly to author a"
                            + " supported AGILE package instead.")),
        ExcelOoxmlSecurityPoiBridge.toWriteHash(sourceEncryption.hashAlgorithm())
            .orElseThrow(
                () ->
                    new IllegalArgumentException(
                        "Source workbook encryption hash "
                            + sourceEncryption.hashAlgorithm()
                            + " is readable but not auto-preservable on write;"
                            + " set persistence.security.encryption explicitly to author a"
                            + " supported AGILE package instead.")));
  }
}
