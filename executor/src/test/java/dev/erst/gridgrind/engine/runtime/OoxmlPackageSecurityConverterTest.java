package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceEncryption;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceSignature;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests the protocol-to-engine OOXML security conversion through prepared file capabilities. */
class OoxmlPackageSecurityConverterTest {
  @TempDir Path temporaryDirectory;

  @Test
  void convertsPresentSecuritySettingsIntoEngineOptions() throws Exception {
    Path signingMaterial =
        Files.write(temporaryDirectory.resolve("signing-material.p12"), new byte[] {1, 2});
    OoxmlPersistenceSecurityInput input =
        new OoxmlPersistenceSecurityInput(
            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput.Encrypt(
                new OoxmlEncryptionInput(
                    new dev.erst.gridgrind.contract.dto.SecretReference("persist-pass"),
                    ExcelOoxmlWriteCipher.AES_192,
                    ExcelOoxmlWriteHash.SHA_384)),
            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.Sign(
                new OoxmlSignatureInput(
                    signingMaterial.getFileName().toString(),
                    new dev.erst.gridgrind.contract.dto.SecretReference("keystore-password"),
                    Optional.of(
                        new dev.erst.gridgrind.contract.dto.SecretReference("key-password")),
                    Optional.of("gridgrind-signing"),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA512,
                    Optional.of("GridGrind signing test"))));

    try (var prepared =
        ExecutionInputBindingsFixtureSupport.preparedBindings(
            temporaryDirectory,
            List.of(signingMaterial),
            java.util.Map.of(
                new dev.erst.gridgrind.contract.dto.SecretReference("persist-pass"),
                "persist-pass",
                new dev.erst.gridgrind.contract.dto.SecretReference("keystore-password"),
                "keystore-pass",
                new dev.erst.gridgrind.contract.dto.SecretReference("key-password"),
                "key-pass",
                new dev.erst.gridgrind.contract.dto.SecretReference("source-pass"),
                "source-pass"))) {
      assertEquals(
          "source-pass",
          assertInstanceOf(
                  ExcelOoxmlOpenOptions.Encrypted.class,
                  OoxmlPackageSecurityConverter.toExcelOpenOptions(
                      new OoxmlOpenSecurityInput(
                          Optional.of(
                              new dev.erst.gridgrind.contract.dto.SecretReference("source-pass"))),
                      prepared.bindings()))
              .password());
      ExcelOoxmlPersistenceOptions options =
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input, prepared.bindings());
      var encryption =
          assertInstanceOf(ExcelOoxmlPersistenceEncryption.Encrypt.class, options.encryption())
              .options();
      assertEquals("persist-pass", encryption.password());
      assertEquals(ExcelOoxmlWriteCipher.AES_192, encryption.cipher());
      assertEquals(ExcelOoxmlWriteHash.SHA_384, encryption.hash());
      var signature =
          assertInstanceOf(ExcelOoxmlPersistenceSignature.Sign.class, options.signature())
              .options();
      assertArrayEquals(new byte[] {1, 2}, Files.readAllBytes(signature.pkcs12Path()));
      assertEquals("key-pass", signature.keyPassword());
      assertEquals("gridgrind-signing", signature.alias());
      assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA512, signature.digestAlgorithm());
    }
  }

  @Test
  void convertsExplicitNoneSecuritySettingsIntoPlaintextEngineOptions() throws Exception {
    OoxmlPersistenceSecurityInput explicitNone =
        new OoxmlPersistenceSecurityInput(
            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput.None(),
            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.None());

    assertInstanceOf(
        ExcelOoxmlOpenOptions.Unencrypted.class,
        OoxmlPackageSecurityConverter.toExcelOpenOptions(null));
    assertInstanceOf(
        ExcelOoxmlOpenOptions.Unencrypted.class,
        OoxmlPackageSecurityConverter.toExcelOpenOptions(
            new OoxmlOpenSecurityInput(Optional.empty())));
    assertThrows(
        IllegalStateException.class,
        () ->
            OoxmlPackageSecurityConverter.toExcelOpenOptions(
                new OoxmlOpenSecurityInput(
                    Optional.of(
                        new dev.erst.gridgrind.contract.dto.SecretReference("source-pass")))));
    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(temporaryDirectory)) {
      assertInstanceOf(
          ExcelOoxmlOpenOptions.Unencrypted.class,
          OoxmlPackageSecurityConverter.toExcelOpenOptions(
              new OoxmlOpenSecurityInput(Optional.empty()), prepared.bindings()));
      assertTrue(
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(null, prepared.bindings())
              .writesPlaintextUnsigned());
      assertInstanceOf(
          ExcelOoxmlPersistenceSignature.Unsigned.class,
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(explicitNone, prepared.bindings())
              .signature());
    }
  }

  @Test
  void materializesRelativeSigningMaterialThroughThePreparedExecutionRoot() throws Exception {
    Path keys = Files.createDirectory(temporaryDirectory.resolve("keys"));
    Path signingMaterial = Files.write(keys.resolve("signing-material.p12"), new byte[] {3, 4, 5});
    OoxmlPersistenceSecurityInput input =
        new OoxmlPersistenceSecurityInput(
            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput.None(),
            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.Sign(
                new OoxmlSignatureInput(
                    "keys/signing-material.p12",
                    new dev.erst.gridgrind.contract.dto.SecretReference("keystore-password"),
                    Optional.of(
                        new dev.erst.gridgrind.contract.dto.SecretReference("key-password")),
                    Optional.of("gridgrind-signing"),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                    Optional.of("GridGrind signing test"))));

    try (var prepared =
        ExecutionInputBindingsFixtureSupport.preparedBindings(
            temporaryDirectory,
            List.of(signingMaterial),
            java.util.Map.of(
                new dev.erst.gridgrind.contract.dto.SecretReference("keystore-password"),
                "keystore-pass",
                new dev.erst.gridgrind.contract.dto.SecretReference("key-password"),
                "key-pass"))) {
      Path materialized =
          assertInstanceOf(
                  ExcelOoxmlPersistenceSignature.Sign.class,
                  OoxmlPackageSecurityConverter.toExcelPersistenceOptions(
                          input, prepared.bindings())
                      .signature())
              .options()
              .pkcs12Path();
      assertArrayEquals(Files.readAllBytes(signingMaterial), Files.readAllBytes(materialized));
      assertTrue(materialized.startsWith(temporaryDirectory.resolve(".gridgrind/tmp")));
    }
  }

  @Test
  void usesTheKeystorePasswordForSigningWhenNoDistinctKeyPasswordWasRequested() throws Exception {
    Path signingMaterial =
        Files.write(temporaryDirectory.resolve("signing-material.p12"), new byte[] {1});
    OoxmlPersistenceSecurityInput input =
        new OoxmlPersistenceSecurityInput(
            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput.None(),
            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.Sign(
                OoxmlSignatureInput.sameKeyPassword(
                    signingMaterial.getFileName().toString(),
                    new dev.erst.gridgrind.contract.dto.SecretReference("keystore-password"),
                    Optional.empty(),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                    Optional.empty())));

    try (var prepared =
        ExecutionInputBindingsFixtureSupport.preparedBindings(
            temporaryDirectory,
            List.of(signingMaterial),
            java.util.Map.of(
                new dev.erst.gridgrind.contract.dto.SecretReference("keystore-password"),
                "keystore-pass"))) {
      assertEquals(
          "keystore-pass",
          assertInstanceOf(
                  ExcelOoxmlPersistenceSignature.Sign.class,
                  OoxmlPackageSecurityConverter.toExcelPersistenceOptions(
                          input, prepared.bindings())
                      .signature())
              .options()
              .keyPassword());
    }
  }
}
