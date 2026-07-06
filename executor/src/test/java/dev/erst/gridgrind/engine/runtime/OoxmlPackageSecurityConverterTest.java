package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlOpenOptions;
import java.nio.file.Path;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for the protocol-to-engine OOXML package-security conversion seam. */
class OoxmlPackageSecurityConverterTest {
  @Test
  void convertsPresentSecuritySettingsIntoEngineOptions() {
    Path workingDirectory = Path.of("/tmp/gridgrind-ooxml-security");
    OoxmlPersistenceSecurityInput input =
        new OoxmlPersistenceSecurityInput(
            new OoxmlEncryptionInput(
                "persist-pass", ExcelOoxmlWriteCipher.AES_192, ExcelOoxmlWriteHash.SHA_384),
            new OoxmlSignatureInput(
                "/tmp/signing-material.p12",
                "keystore-pass",
                "key-pass",
                Optional.of("gridgrind-signing"),
                ExcelOoxmlSignatureDigestAlgorithm.SHA512,
                Optional.of("GridGrind signing test")));

    assertEquals(
        "source-pass",
        assertInstanceOf(
                ExcelOoxmlOpenOptions.Encrypted.class,
                OoxmlPackageSecurityConverter.toExcelOpenOptions(
                    new OoxmlOpenSecurityInput(Optional.of("source-pass"))))
            .password());
    assertEquals(
        "persist-pass",
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input, workingDirectory)
            .encryption()
            .orElseThrow()
            .password());
    assertEquals(
        ExcelOoxmlWriteCipher.AES_192,
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input, workingDirectory)
            .encryption()
            .orElseThrow()
            .cipher());
    assertEquals(
        ExcelOoxmlWriteHash.SHA_384,
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input, workingDirectory)
            .encryption()
            .orElseThrow()
            .hash());
    assertEquals(
        Path.of("/tmp/signing-material.p12"),
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input, workingDirectory)
            .signature()
            .orElseThrow()
            .pkcs12Path());
    assertEquals(
        "key-pass",
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input, workingDirectory)
            .signature()
            .orElseThrow()
            .keyPassword());
    assertEquals(
        "gridgrind-signing",
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input, workingDirectory)
            .signature()
            .orElseThrow()
            .alias());
    assertEquals(
        ExcelOoxmlSignatureDigestAlgorithm.SHA512,
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input, workingDirectory)
            .signature()
            .orElseThrow()
            .digestAlgorithm());
  }

  @Test
  void convertsMissingSecuritySettingsIntoEmptyEngineOptions() {
    Path workingDirectory = Path.of("/tmp/gridgrind-ooxml-security");
    OoxmlPersistenceSecurityInput encryptionOnly =
        new OoxmlPersistenceSecurityInput(
            new OoxmlEncryptionInput(
                "persist-pass", ExcelOoxmlWriteCipher.AES_256, ExcelOoxmlWriteHash.SHA_512),
            null);

    assertInstanceOf(
        ExcelOoxmlOpenOptions.Unencrypted.class,
        OoxmlPackageSecurityConverter.toExcelOpenOptions(null));
    assertTrue(
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(null, workingDirectory).isEmpty());
    assertTrue(
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(encryptionOnly, workingDirectory)
            .signature()
            .isEmpty());
  }

  @Test
  void rootsRelativeSigningMaterialPathsInTheProvidedWorkingDirectory() {
    OoxmlPersistenceSecurityInput input =
        new OoxmlPersistenceSecurityInput(
            null,
            new OoxmlSignatureInput(
                "keys/signing-material.p12",
                "keystore-pass",
                "key-pass",
                Optional.of("gridgrind-signing"),
                ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                Optional.of("GridGrind signing test")));
    Path workingDirectory = Path.of("/tmp/gridgrind-request-bundle");

    assertEquals(
        workingDirectory.resolve("keys/signing-material.p12").normalize(),
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input, workingDirectory)
            .signature()
            .orElseThrow()
            .pkcs12Path());
  }
}
