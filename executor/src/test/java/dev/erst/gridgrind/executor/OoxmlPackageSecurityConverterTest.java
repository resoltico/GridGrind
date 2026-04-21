package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.excel.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.ExcelOoxmlSignatureDigestAlgorithm;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Tests for the protocol-to-engine OOXML package-security conversion seam. */
class OoxmlPackageSecurityConverterTest {
  @Test
  void convertsPresentSecuritySettingsIntoEngineOptions() {
    OoxmlPersistenceSecurityInput input =
        new OoxmlPersistenceSecurityInput(
            new OoxmlEncryptionInput("persist-pass", ExcelOoxmlEncryptionMode.STANDARD),
            new OoxmlSignatureInput(
                "/tmp/signing-material.p12",
                "keystore-pass",
                "key-pass",
                "gridgrind-signing",
                ExcelOoxmlSignatureDigestAlgorithm.SHA512,
                "GridGrind signing test"));

    assertEquals(
        "source-pass",
        assertInstanceOf(
                ExcelOoxmlOpenOptions.Encrypted.class,
                OoxmlPackageSecurityConverter.toExcelOpenOptions(
                    new OoxmlOpenSecurityInput("source-pass")))
            .password());
    assertEquals(
        "persist-pass",
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input).encryption().password());
    assertEquals(
        ExcelOoxmlEncryptionMode.STANDARD,
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input).encryption().mode());
    assertEquals(
        Path.of("/tmp/signing-material.p12"),
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input).signature().pkcs12Path());
    assertEquals(
        "key-pass",
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input).signature().keyPassword());
    assertEquals(
        "gridgrind-signing",
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input).signature().alias());
    assertEquals(
        ExcelOoxmlSignatureDigestAlgorithm.SHA512,
        OoxmlPackageSecurityConverter.toExcelPersistenceOptions(input)
            .signature()
            .digestAlgorithm());
  }

  @Test
  void convertsMissingSecuritySettingsIntoEmptyEngineOptions() {
    OoxmlPersistenceSecurityInput encryptionOnly =
        new OoxmlPersistenceSecurityInput(
            new OoxmlEncryptionInput("persist-pass", ExcelOoxmlEncryptionMode.AGILE), null);

    assertInstanceOf(
        ExcelOoxmlOpenOptions.Unencrypted.class,
        OoxmlPackageSecurityConverter.toExcelOpenOptions(null));
    assertTrue(OoxmlPackageSecurityConverter.toExcelPersistenceOptions(null).isEmpty());
    assertNull(OoxmlPackageSecurityConverter.toExcelPersistenceOptions(encryptionOnly).signature());
  }
}
