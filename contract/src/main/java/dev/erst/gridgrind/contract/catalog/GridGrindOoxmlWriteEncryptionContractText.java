package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteHash;
import java.util.Arrays;

/** Stable public wording for the OOXML write-encryption contract and operator-facing limit. */
public final class GridGrindOoxmlWriteEncryptionContractText {
  private GridGrindOoxmlWriteEncryptionContractText() {}

  /** One stable catalog summary for `OoxmlEncryptionInput`. */
  public static String inputSummary() { // LIM-038
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
        + " (LIM-038).";
  }

  /** One stable operator-facing summary of the OOXML write-encryption limitation. */
  public static String limitSummary() { // LIM-038
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
