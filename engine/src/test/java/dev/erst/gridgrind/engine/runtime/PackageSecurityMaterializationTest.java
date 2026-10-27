package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.security.OoxmlEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureSpec;
import dev.erst.gridgrind.core.model.security.ResolvedOpenSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceEncryption;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSignature;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests the protocol-to-engine OOXML security conversion through prepared file capabilities. */
class PackageSecurityMaterializationTest {
  @TempDir Path temporaryDirectory;

  @Test
  void convertsPresentSecuritySettingsIntoEngineOptions() throws Exception {
    Path signingMaterial =
        Files.write(temporaryDirectory.resolve("signing-material.p12"), new byte[] {1, 2});
    OoxmlPersistenceSecuritySpec input =
        new OoxmlPersistenceSecuritySpec(
            new dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.Encrypt(
                new OoxmlEncryptionSpec(
                    new dev.erst.gridgrind.core.source.SecretReference("persist-pass"),
                    ExcelOoxmlWriteCipher.AES_192,
                    ExcelOoxmlWriteHash.SHA_384)),
            new dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec.Sign(
                new OoxmlSignatureSpec(
                    signingMaterial.getFileName().toString(),
                    new dev.erst.gridgrind.core.source.SecretReference("keystore-password"),
                    Optional.of(new dev.erst.gridgrind.core.source.SecretReference("key-password")),
                    Optional.of("gridgrind-signing"),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA512,
                    Optional.of("GridGrind signing test"))));

    try (var prepared =
        ExecutionInputBindingsFixtureSupport.preparedBindings(
            temporaryDirectory,
            List.of(signingMaterial),
            java.util.Map.of(
                new dev.erst.gridgrind.core.source.SecretReference("persist-pass"),
                "persist-pass",
                new dev.erst.gridgrind.core.source.SecretReference("keystore-password"),
                "keystore-pass",
                new dev.erst.gridgrind.core.source.SecretReference("key-password"),
                "key-pass",
                new dev.erst.gridgrind.core.source.SecretReference("source-pass"),
                "source-pass"))) {
      assertEquals(
          "source-pass",
          assertInstanceOf(
                  ResolvedOpenSecurity.Encrypted.class,
                  ExecutionRequestPaths.materializeOpenSecurity(
                      new OoxmlOpenSecuritySpec(
                          Optional.of(
                              new dev.erst.gridgrind.core.source.SecretReference("source-pass"))),
                      prepared.bindings()))
              .password());
      ResolvedPersistenceSecurity options =
          ExecutionRequestPaths.materializePersistenceSecurity(input, prepared.bindings());
      var encryption =
          assertInstanceOf(ResolvedPersistenceEncryption.Encrypt.class, options.encryption())
              .options();
      assertEquals("persist-pass", encryption.password());
      assertEquals(ExcelOoxmlWriteCipher.AES_192, encryption.cipher());
      assertEquals(ExcelOoxmlWriteHash.SHA_384, encryption.hash());
      var signature =
          assertInstanceOf(ResolvedPersistenceSignature.Sign.class, options.signature()).options();
      assertArrayEquals(new byte[] {1, 2}, Files.readAllBytes(signature.pkcs12Path()));
      assertEquals("key-pass", signature.keyPassword());
      assertEquals(Optional.of("gridgrind-signing"), signature.alias());
      assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA512, signature.digestAlgorithm());
    }
  }

  @Test
  void materializesExplicitNoneSecurityAsPlaintextUnsigned() throws Exception {
    OoxmlPersistenceSecuritySpec explicitNone =
        new OoxmlPersistenceSecuritySpec(
            new dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.None(),
            new dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec.None());

    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(temporaryDirectory)) {
      assertInstanceOf(
          ResolvedOpenSecurity.Unencrypted.class,
          ExecutionRequestPaths.materializeOpenSecurity(null, prepared.bindings()));
      assertInstanceOf(
          ResolvedOpenSecurity.Unencrypted.class,
          ExecutionRequestPaths.materializeOpenSecurity(
              new OoxmlOpenSecuritySpec(Optional.empty()), prepared.bindings()));
      assertTrue(
          ExecutionRequestPaths.materializePersistenceSecurity(null, prepared.bindings())
              .writesPlaintextUnsigned());
      assertInstanceOf(
          OoxmlPersistenceSignatureSpec.None.class,
          ExecutionRequestPaths.materializePersistenceSecurity(explicitNone, prepared.bindings())
              .signature());
    }
  }

  @Test
  void materializesRelativeSigningMaterialThroughThePreparedExecutionRoot() throws Exception {
    Path keys = Files.createDirectory(temporaryDirectory.resolve("keys"));
    Path signingMaterial = Files.write(keys.resolve("signing-material.p12"), new byte[] {3, 4, 5});
    OoxmlPersistenceSecuritySpec input =
        new OoxmlPersistenceSecuritySpec(
            new dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.None(),
            new dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec.Sign(
                new OoxmlSignatureSpec(
                    "keys/signing-material.p12",
                    new dev.erst.gridgrind.core.source.SecretReference("keystore-password"),
                    Optional.of(new dev.erst.gridgrind.core.source.SecretReference("key-password")),
                    Optional.of("gridgrind-signing"),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                    Optional.of("GridGrind signing test"))));

    try (var prepared =
        ExecutionInputBindingsFixtureSupport.preparedBindings(
            temporaryDirectory,
            List.of(signingMaterial),
            java.util.Map.of(
                new dev.erst.gridgrind.core.source.SecretReference("keystore-password"),
                "keystore-pass",
                new dev.erst.gridgrind.core.source.SecretReference("key-password"),
                "key-pass"))) {
      Path materialized =
          assertInstanceOf(
                  ResolvedPersistenceSignature.Sign.class,
                  ExecutionRequestPaths.materializePersistenceSecurity(input, prepared.bindings())
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
    OoxmlPersistenceSecuritySpec input =
        new OoxmlPersistenceSecuritySpec(
            new dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.None(),
            new dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec.Sign(
                OoxmlSignatureSpec.sameKeyPassword(
                    signingMaterial.getFileName().toString(),
                    new dev.erst.gridgrind.core.source.SecretReference("keystore-password"),
                    Optional.empty(),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                    Optional.empty())));

    try (var prepared =
        ExecutionInputBindingsFixtureSupport.preparedBindings(
            temporaryDirectory,
            List.of(signingMaterial),
            java.util.Map.of(
                new dev.erst.gridgrind.core.source.SecretReference("keystore-password"),
                "keystore-pass"))) {
      assertEquals(
          "keystore-pass",
          assertInstanceOf(
                  ResolvedPersistenceSignature.Sign.class,
                  ExecutionRequestPaths.materializePersistenceSecurity(input, prepared.bindings())
                      .signature())
              .options()
              .keyPassword());
    }
  }
}
