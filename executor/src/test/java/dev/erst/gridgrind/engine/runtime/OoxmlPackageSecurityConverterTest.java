package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
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
                    "persist-pass", ExcelOoxmlWriteCipher.AES_192, ExcelOoxmlWriteHash.SHA_384)),
            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.Sign(
                new OoxmlSignatureInput(
                    signingMaterial.getFileName().toString(),
                    "keystore-pass",
                    "key-pass",
                    Optional.of("gridgrind-signing"),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA512,
                    Optional.of("GridGrind signing test"))));

    assertEquals(
        "source-pass",
        assertInstanceOf(
                ExcelOoxmlOpenOptions.Encrypted.class,
                OoxmlPackageSecurityConverter.toExcelOpenOptions(
                    new OoxmlOpenSecurityInput(Optional.of("source-pass"))))
            .password());
    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(temporaryDirectory)) {
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
    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(temporaryDirectory)) {
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
                    "keystore-pass",
                    "key-pass",
                    Optional.of("gridgrind-signing"),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                    Optional.of("GridGrind signing test"))));

    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(temporaryDirectory)) {
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
}
