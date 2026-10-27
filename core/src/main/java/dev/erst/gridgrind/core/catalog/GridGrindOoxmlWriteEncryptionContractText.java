package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import java.util.Arrays;

/** Stable public wording for the OOXML write-encryption contract and operator-facing limit. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_OOXML_WRITE_ENCRYPTION_IS_AGILE_ONLY_AND_STRONG_ONLY
})
public final class GridGrindOoxmlWriteEncryptionContractText {
  private GridGrindOoxmlWriteEncryptionContractText() {}

  /** One stable catalog summary for `OoxmlEncryptionSpec`. */
  public static String inputSummary() {
    return "OOXML package-encryption settings for workbook persistence."
        + " GridGrind writes AGILE packages only."
        + " cipher defaults to "
        + ExcelOoxmlWriteCipher.AES_256
        + " and hash defaults to "
        + ExcelOoxmlWriteHash.SHA_512
        + " when omitted."
        + " Supported ciphers: "
        + enumValuePhrase(ExcelOoxmlWriteCipher.values())
        + ". Supported hashes: "
        + enumValuePhrase(ExcelOoxmlWriteHash.values())
        + " ("
        + dev.erst.gridgrind.core.limits.Limit
            .ID_OOXML_WRITE_ENCRYPTION_IS_AGILE_ONLY_AND_STRONG_ONLY
        + ").";
  }

  /** One stable operator-facing summary of the OOXML write-encryption limitation. */
  public static String limitSummary() {
    return "writes AGILE packages only; cipher defaults to "
        + ExcelOoxmlWriteCipher.AES_256
        + ", hash defaults to "
        + ExcelOoxmlWriteHash.SHA_512
        + ", supported ciphers are "
        + enumValuePhrase(ExcelOoxmlWriteCipher.values())
        + ", supported hashes are "
        + enumValuePhrase(ExcelOoxmlWriteHash.values())
        + ", and STANDARD remains readable on inspection but is not authorable.";
  }

  private static String enumValuePhrase(Enum<?>[] values) {
    return GridGrindContractText.humanJoin(
        Arrays.stream(values).map(Enum::name).map(String::trim).toList());
  }
}
