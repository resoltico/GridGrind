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
    return toExcelPersistenceOptions(input, Path.of(""));
  }

  static ExcelOoxmlPersistenceOptions toExcelPersistenceOptions(
      OoxmlPersistenceSecurityInput input, Path workingDirectory) {
    if (input == null) {
      return new ExcelOoxmlPersistenceOptions(null, null);
    }
    return new ExcelOoxmlPersistenceOptions(
        toExcelEncryptionOptions(input.encryption()),
        toExcelSignatureOptions(input.signature(), workingDirectory));
  }

  private static ExcelOoxmlEncryptionOptions toExcelEncryptionOptions(OoxmlEncryptionInput input) {
    return input == null ? null : new ExcelOoxmlEncryptionOptions(input.password(), input.mode());
  }

  private static ExcelOoxmlSignatureOptions toExcelSignatureOptions(
      OoxmlSignatureInput input, Path workingDirectory) {
    return input == null
        ? null
        : new ExcelOoxmlSignatureOptions(
            ExecutionRequestPaths.normalizePath(input.pkcs12Path(), workingDirectory),
            input.keystorePassword(),
            input.keyPassword(),
            input.alias(),
            input.digestAlgorithm(),
            input.description());
  }
}
