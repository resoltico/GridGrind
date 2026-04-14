package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.junit.jupiter.api.Test;

/** Direct validation tests for OOXML package-security engine value types. */
class ExcelOoxmlPackageSecurityTypesTest {
  @Test
  void openOptionsEncryptionOptionsAndSignatureOptionsNormalizeInputs() {
    ExcelOoxmlOpenOptions openOptions = new ExcelOoxmlOpenOptions("secret");
    ExcelOoxmlEncryptionOptions encryptionOptions = new ExcelOoxmlEncryptionOptions("secret", null);
    ExcelOoxmlSignatureOptions signatureOptions =
        new ExcelOoxmlSignatureOptions(
            Path.of("/tmp/signing-material.p12"), "store-pass", null, null, null, null);

    assertEquals("secret", openOptions.password());
    assertEquals(ExcelOoxmlEncryptionMode.AGILE, encryptionOptions.mode());
    assertEquals("store-pass", signatureOptions.keyPassword());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signatureOptions.digestAlgorithm());
    assertNull(signatureOptions.alias());
    assertNull(signatureOptions.description());

    assertThrows(IllegalArgumentException.class, () -> new ExcelOoxmlOpenOptions(" "));
    assertThrows(IllegalArgumentException.class, () -> new ExcelOoxmlEncryptionOptions(" ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlSignatureOptions(
                Path.of("/tmp/signing-material.p12"), "store-pass", " ", null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlSignatureOptions(
                Path.of("/tmp/signing-material.p12"), "store-pass", null, " ", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlSignatureOptions(
                Path.of("/tmp/signing-material.p12"), "store-pass", null, null, null, " "));
  }

  @Test
  void encryptionAndSignatureSnapshotsValidateSecurityFacts() {
    ExcelOoxmlEncryptionSnapshot encryption =
        new ExcelOoxmlEncryptionSnapshot(
            true,
            ExcelOoxmlEncryptionMode.AGILE,
            "AES256",
            "SHA512",
            "ChainingModeCBC",
            256,
            16,
            100_000);
    ExcelOoxmlSignatureSnapshot signature =
        new ExcelOoxmlSignatureSnapshot(
            "/_xmlsignatures/sig1.xml",
            "CN=GridGrind Signing Test",
            "CN=GridGrind Signing Test",
            "01AB",
            ExcelOoxmlSignatureState.VALID);
    ExcelOoxmlPackageSecuritySnapshot packageSecurity =
        new ExcelOoxmlPackageSecuritySnapshot(encryption, List.of(signature));

    assertTrue(packageSecurity.isSecure());
    assertEquals(
        ExcelOoxmlSignatureState.INVALIDATED_BY_MUTATION,
        packageSecurity.afterMutation().signatures().getFirst().state());
    assertFalse(ExcelOoxmlPackageSecuritySnapshot.none().isSecure());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                false, ExcelOoxmlEncryptionMode.AGILE, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true, ExcelOoxmlEncryptionMode.AGILE, "AES256", "SHA512", "CBC", 0, 16, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlSignatureSnapshot(
                " ", "CN=GridGrind Signing Test", null, null, ExcelOoxmlSignatureState.VALID));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelOoxmlPackageSecuritySnapshot(encryption, Arrays.asList(signature, null)));

    assertEquals(
        ExcelOoxmlSignatureState.INVALIDATED_BY_MUTATION,
        ExcelOoxmlSignatureState.VALID.afterMutation());
    assertEquals(
        ExcelOoxmlSignatureState.INVALID, ExcelOoxmlSignatureState.INVALID.afterMutation());
  }

  @Test
  void coverageBranchesForSecurityTypesAndExceptionsStayExplicit() throws IOException {
    assertNull(new ExcelOoxmlOpenOptions(null).password());
    assertEquals(
        ExcelOoxmlEncryptionMode.AGILE, ExcelOoxmlEncryptionMode.fromPoi(EncryptionMode.agile));
    assertEquals(
        ExcelOoxmlEncryptionMode.STANDARD,
        ExcelOoxmlEncryptionMode.fromPoi(EncryptionMode.standard));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelOoxmlEncryptionMode.fromPoi(EncryptionMode.binaryRC4));

    ExcelOoxmlEncryptionSnapshot none = ExcelOoxmlEncryptionSnapshot.none();
    assertFalse(none.encrypted());
    assertNull(none.mode());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true, ExcelOoxmlEncryptionMode.AGILE, null, "SHA512", "CBC", 256, 16, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true, ExcelOoxmlEncryptionMode.AGILE, "AES256", " ", "CBC", 256, 16, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true, ExcelOoxmlEncryptionMode.AGILE, "AES256", "SHA512", " ", 256, 16, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true, ExcelOoxmlEncryptionMode.AGILE, "AES256", "SHA512", "CBC", 256, 0, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true, ExcelOoxmlEncryptionMode.AGILE, "AES256", "SHA512", "CBC", 256, 16, -1));
    // !encrypted compound-OR: each condition needs to be the "first true" to cover its branch.
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(false, null, "AES256", null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(false, null, null, "SHA512", null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                false, null, null, null, "ChainingModeCBC", null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelOoxmlEncryptionSnapshot(false, null, null, null, null, 256, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelOoxmlEncryptionSnapshot(false, null, null, null, null, null, 16, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelOoxmlEncryptionSnapshot(false, null, null, null, null, null, null, 100_000));
    // encrypted path: null keyBits, null blockSize, null spinCount.
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true, ExcelOoxmlEncryptionMode.AGILE, "AES256", "SHA512", "CBC", null, 16, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true, ExcelOoxmlEncryptionMode.AGILE, "AES256", "SHA512", "CBC", 256, null, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true, ExcelOoxmlEncryptionMode.AGILE, "AES256", "SHA512", "CBC", 256, 16, null));

    ExcelOoxmlPackageSecuritySnapshot plainSecurity = ExcelOoxmlPackageSecuritySnapshot.none();
    assertSame(plainSecurity, plainSecurity.afterMutation());

    Path workbookPath = Path.of("/tmp/security.xlsx");
    WorkbookPasswordRequiredException passwordRequired =
        new WorkbookPasswordRequiredException(workbookPath);
    assertEquals(workbookPath, passwordRequired.workbookPath());
    assertTrue(passwordRequired.getMessage().contains("source.security.password"));

    InvalidWorkbookPasswordException invalidPassword =
        new InvalidWorkbookPasswordException(workbookPath);
    assertEquals(workbookPath, invalidPassword.workbookPath());
    assertTrue(invalidPassword.getMessage().contains("did not unlock the workbook"));

    IllegalStateException cause = new IllegalStateException("boom");
    InvalidSigningConfigurationException invalidSigningConfiguration =
        new InvalidSigningConfigurationException("signing problem", cause);
    assertEquals("signing problem", invalidSigningConfiguration.getMessage());
    assertSame(cause, invalidSigningConfiguration.getCause());
    assertEquals(
        "simple signing problem",
        new InvalidSigningConfigurationException("simple signing problem").getMessage());

    WorkbookSecurityException securityException = new WorkbookSecurityException("security problem");
    assertEquals("security problem", securityException.getMessage());
    assertNull(securityException.getCause());

    WorkbookSecurityException wrappedSecurityException =
        new WorkbookSecurityException("wrapped security problem", cause);
    assertEquals("wrapped security problem", wrappedSecurityException.getMessage());
    assertSame(cause, wrappedSecurityException.getCause());
  }
}
