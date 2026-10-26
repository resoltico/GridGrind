package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.security.OoxmlEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureSpec;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceEncryption;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceSignature;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureOptions;
import java.io.IOException;
import org.jspecify.annotations.Nullable;

/** Converts protocol OOXML package-security DTOs into engine-owned security option shapes. */
final class OoxmlPackageSecurityConverter {
  private OoxmlPackageSecurityConverter() {}

  static ExcelOoxmlOpenOptions toExcelOpenOptions(
      @Nullable OoxmlOpenSecuritySpec input, ExecutionInputBindings bindings) {
    return input == null || input.passwordRef().isEmpty()
        ? new ExcelOoxmlOpenOptions.Unencrypted()
        : new ExcelOoxmlOpenOptions.Encrypted(
            bindings.resolveSecret(input.passwordRef().orElseThrow()));
  }

  static ExcelOoxmlOpenOptions toExcelOpenOptions(@Nullable OoxmlOpenSecuritySpec input) {
    if (input == null || input.passwordRef().isEmpty()) {
      return new ExcelOoxmlOpenOptions.Unencrypted();
    }
    throw new IllegalStateException("encrypted open options require execution bindings");
  }

  static ExcelOoxmlPersistenceOptions toExcelPersistenceOptions(
      @Nullable OoxmlPersistenceSecuritySpec input, ExecutionInputBindings bindings)
      throws IOException {
    if (input == null) {
      return ExcelOoxmlPersistenceOptions.none();
    }
    return new ExcelOoxmlPersistenceOptions(
        encryptionPolicy(input.encryption(), bindings),
        signaturePolicy(input.signature(), bindings));
  }

  private static ExcelOoxmlPersistenceEncryption encryptionPolicy(
      OoxmlPersistenceEncryptionSpec input, ExecutionInputBindings bindings) {
    return switch (input) {
      case OoxmlPersistenceEncryptionSpec.None _ -> new ExcelOoxmlPersistenceEncryption.Plaintext();
      case OoxmlPersistenceEncryptionSpec.Encrypt encrypt ->
          new ExcelOoxmlPersistenceEncryption.Encrypt(
              toExcelEncryptionOptions(encrypt.encryption(), bindings));
      case OoxmlPersistenceEncryptionSpec.PreserveSource _ ->
          new ExcelOoxmlPersistenceEncryption.PreserveSource();
    };
  }

  private static ExcelOoxmlPersistenceSignature signaturePolicy(
      OoxmlPersistenceSignatureSpec input, ExecutionInputBindings bindings) throws IOException {
    return switch (input) {
      case OoxmlPersistenceSignatureSpec.None _ -> new ExcelOoxmlPersistenceSignature.Unsigned();
      case OoxmlPersistenceSignatureSpec.Sign sign ->
          new ExcelOoxmlPersistenceSignature.Sign(
              toExcelSignatureOptions(sign.signature(), bindings));
    };
  }

  private static ExcelOoxmlEncryptionOptions toExcelEncryptionOptions(
      OoxmlEncryptionSpec input, ExecutionInputBindings bindings) {
    return new ExcelOoxmlEncryptionOptions(
        bindings.resolveSecret(input.passwordRef()), input.cipher(), input.hash());
  }

  private static ExcelOoxmlSignatureOptions toExcelSignatureOptions(
      OoxmlSignatureSpec input, ExecutionInputBindings bindings) throws IOException {
    try {
      return new ExcelOoxmlSignatureOptions(
          bindings
              .requestPathAccess()
              .materializeRead(
                  input.pkcs12Path(),
                  "persistence.security.signature.signature.pkcs12Path",
                  "gridgrind-signing-material-",
                  ".p12"),
          bindings.resolveSecret(input.keystorePasswordRef()),
          input
              .keyPasswordRef()
              .map(bindings::resolveSecret)
              .orElseGet(() -> bindings.resolveSecret(input.keystorePasswordRef())),
          input.alias().orElse(null),
          input.digestAlgorithm(),
          input.description().orElse(null));
    } catch (java.nio.file.NoSuchFileException exception) {
      throw new dev.erst.gridgrind.excel.InvalidSigningConfigurationException(
          "Signing material does not exist: " + input.pkcs12Path(), exception);
    }
  }
}
