package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureOptions;
import java.io.IOException;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Converts protocol OOXML package-security DTOs into engine-owned security option shapes. */
final class OoxmlPackageSecurityConverter {
  private OoxmlPackageSecurityConverter() {}

  static ExcelOoxmlOpenOptions toExcelOpenOptions(@Nullable OoxmlOpenSecurityInput input) {
    return input == null || input.password().isEmpty()
        ? new ExcelOoxmlOpenOptions.Unencrypted()
        : new ExcelOoxmlOpenOptions.Encrypted(input.password().orElseThrow());
  }

  static ExcelOoxmlPersistenceOptions toExcelPersistenceOptions(
      @Nullable OoxmlPersistenceSecurityInput input, ExecutionInputBindings bindings)
      throws IOException {
    if (input == null) {
      return ExcelOoxmlPersistenceOptions.none();
    }
    return new ExcelOoxmlPersistenceOptions(
        Optional.ofNullable(toExcelEncryptionOptions(input.encryption())),
        Optional.ofNullable(toExcelSignatureOptions(input.signature(), bindings)));
  }

  private static @Nullable ExcelOoxmlEncryptionOptions toExcelEncryptionOptions(
      @Nullable OoxmlEncryptionInput input) {
    return input == null
        ? null
        : new ExcelOoxmlEncryptionOptions(input.password(), input.cipher(), input.hash());
  }

  private static @Nullable ExcelOoxmlSignatureOptions toExcelSignatureOptions(
      @Nullable OoxmlSignatureInput input, ExecutionInputBindings bindings) throws IOException {
    if (input == null) {
      return null;
    }
    try {
      return new ExcelOoxmlSignatureOptions(
          bindings
              .requestPathAccess()
              .materializeRead(
                  input.pkcs12Path(),
                  "persistence.security.signature.pkcs12Path",
                  "gridgrind-signing-material-",
                  ".p12"),
          input.keystorePassword(),
          input.keyPassword(),
          input.alias().orElse(null),
          input.digestAlgorithm(),
          input.description().orElse(null));
    } catch (java.nio.file.NoSuchFileException exception) {
      throw new dev.erst.gridgrind.excel.InvalidSigningConfigurationException(
          "Signing material does not exist: " + input.pkcs12Path(), exception);
    }
  }
}
