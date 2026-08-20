package dev.erst.gridgrind.excel.ooxml;

import java.util.Objects;

/** Total OOXML package-security policy applied during workbook persistence. */
public record ExcelOoxmlPersistenceOptions(
    ExcelOoxmlPersistenceEncryption encryption, ExcelOoxmlPersistenceSignature signature) {
  public ExcelOoxmlPersistenceOptions {
    Objects.requireNonNull(encryption, "encryption must not be null");
    Objects.requireNonNull(signature, "signature must not be null");
  }

  /** Returns the explicit plaintext and unsigned output policy. */
  public static ExcelOoxmlPersistenceOptions none() {
    return new ExcelOoxmlPersistenceOptions(
        new ExcelOoxmlPersistenceEncryption.Plaintext(),
        new ExcelOoxmlPersistenceSignature.Unsigned());
  }

  /** Returns whether the policy deliberately writes plaintext without a package signature. */
  public boolean writesPlaintextUnsigned() {
    return encryption instanceof ExcelOoxmlPersistenceEncryption.Plaintext
        && signature instanceof ExcelOoxmlPersistenceSignature.Unsigned;
  }
}
