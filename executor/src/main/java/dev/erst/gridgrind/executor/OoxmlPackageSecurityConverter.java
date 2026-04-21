package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.excel.ExcelOoxmlEncryptionOptions;
import dev.erst.gridgrind.excel.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.ExcelOoxmlSignatureOptions;
import java.nio.file.Path;

/** Converts protocol OOXML package-security DTOs into engine-owned security option shapes. */
final class OoxmlPackageSecurityConverter {
  private OoxmlPackageSecurityConverter() {}

  static ExcelOoxmlOpenOptions toExcelOpenOptions(OoxmlOpenSecurityInput input) {
    return input == null
        ? new ExcelOoxmlOpenOptions.Unencrypted()
        : new ExcelOoxmlOpenOptions.Encrypted(input.password());
  }

  static ExcelOoxmlPersistenceOptions toExcelPersistenceOptions(
      OoxmlPersistenceSecurityInput input) {
    if (input == null) {
      return new ExcelOoxmlPersistenceOptions(null, null);
    }
    return new ExcelOoxmlPersistenceOptions(
        toExcelEncryptionOptions(input.encryption()), toExcelSignatureOptions(input.signature()));
  }

  private static ExcelOoxmlEncryptionOptions toExcelEncryptionOptions(OoxmlEncryptionInput input) {
    return input == null ? null : new ExcelOoxmlEncryptionOptions(input.password(), input.mode());
  }

  private static ExcelOoxmlSignatureOptions toExcelSignatureOptions(OoxmlSignatureInput input) {
    return input == null
        ? null
        : new ExcelOoxmlSignatureOptions(
            Path.of(input.pkcs12Path()),
            input.keystorePassword(),
            input.keyPassword(),
            input.alias(),
            input.digestAlgorithm(),
            input.description());
  }
}
