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
              new ExcelOoxmlEncryptionOptions(
                  sourceEncryptionPassword.orElseThrow(), sourceEncryption.mode()));
    }
    return new ExcelOoxmlPersistenceOptions(encryption, persistenceOptions.signature());
  }
}
