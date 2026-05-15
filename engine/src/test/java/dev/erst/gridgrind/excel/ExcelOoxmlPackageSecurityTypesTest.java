package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionSnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSecurityPoiBridge;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureSnapshot;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.junit.jupiter.api.Test;

/** Direct validation tests for OOXML package-security engine value types. */
class ExcelOoxmlPackageSecurityTypesTest {
  @Test
  void openOptionsEncryptionOptionsAndSignatureOptionsNormalizeInputs() {
    ExcelOoxmlOpenOptions openOptions = new ExcelOoxmlOpenOptions.Encrypted("secret");
    ExcelOoxmlEncryptionOptions encryptionOptions = new ExcelOoxmlEncryptionOptions("secret", null);
    ExcelOoxmlSignatureOptions signatureOptions =
        new ExcelOoxmlSignatureOptions(
            Path.of("/tmp/signing-material.p12"), "store-pass", null, null, null, null);

    assertEquals(
        "secret", assertInstanceOf(ExcelOoxmlOpenOptions.Encrypted.class, openOptions).password());
    assertEquals(ExcelOoxmlEncryptionMode.AGILE, encryptionOptions.mode());
    assertEquals("store-pass", signatureOptions.keyPassword());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signatureOptions.digestAlgorithm());
    assertNull(signatureOptions.alias());
    assertNull(signatureOptions.description());

    assertInstanceOf(
        ExcelOoxmlOpenOptions.Unencrypted.class, new ExcelOoxmlOpenOptions.Unencrypted());
    assertThrows(IllegalArgumentException.class, () -> new ExcelOoxmlOpenOptions.Encrypted(" "));
    assertThrows(NullPointerException.class, () -> new ExcelOoxmlOpenOptions.Encrypted(null));
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
            Optional.of(ExcelOoxmlEncryptionMode.AGILE),
            Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
            Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
            Optional.of(ExcelOoxmlChainingMode.CBC),
            Optional.of(256),
            Optional.of(16),
            Optional.of(100_000));
    ExcelOoxmlSignatureSnapshot signature =
        new ExcelOoxmlSignatureSnapshot(
            "/_xmlsignatures/sig1.xml",
            Optional.of("CN=GridGrind Signing Test"),
            Optional.of("CN=GridGrind Signing Test"),
            Optional.of("01AB"),
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
                false,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                Optional.of(ExcelOoxmlChainingMode.CBC),
                Optional.of(0),
                Optional.of(16),
                Optional.of(1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlSignatureSnapshot(
                " ",
                Optional.of("CN=GridGrind Signing Test"),
                Optional.empty(),
                Optional.empty(),
                ExcelOoxmlSignatureState.VALID));
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
    assertInstanceOf(
        ExcelOoxmlOpenOptions.Unencrypted.class, new ExcelOoxmlOpenOptions.Unencrypted());
    assertEquals(
        EncryptionMode.agile, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlEncryptionMode.AGILE));
    assertEquals(
        EncryptionMode.standard,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlEncryptionMode.STANDARD));
    assertEquals(
        ExcelOoxmlEncryptionMode.AGILE, ExcelOoxmlSecurityPoiBridge.fromPoi(EncryptionMode.agile));
    assertEquals(
        ExcelOoxmlEncryptionMode.STANDARD,
        ExcelOoxmlSecurityPoiBridge.fromPoi(EncryptionMode.standard));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelOoxmlSecurityPoiBridge.fromPoi(EncryptionMode.binaryRC4));
    assertEquals(
        HashAlgorithm.sha256,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA256));
    assertEquals(
        HashAlgorithm.sha384,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA384));
    assertEquals(
        HashAlgorithm.sha512,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA512));

    ExcelOoxmlEncryptionSnapshot none = ExcelOoxmlEncryptionSnapshot.none();
    assertFalse(none.encrypted());
    assertTrue(none.mode().isEmpty());
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                null,
                Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                Optional.of(ExcelOoxmlChainingMode.CBC),
                Optional.of(256),
                Optional.of(16),
                Optional.of(1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                Optional.empty(),
                Optional.of(ExcelOoxmlChainingMode.CBC),
                Optional.of(256),
                Optional.of(16),
                Optional.of(1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                Optional.empty(),
                Optional.of(256),
                Optional.of(16),
                Optional.of(1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                Optional.of(ExcelOoxmlChainingMode.CBC),
                Optional.of(256),
                Optional.of(0),
                Optional.of(1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                Optional.of(ExcelOoxmlChainingMode.CBC),
                Optional.of(256),
                Optional.of(16),
                Optional.of(-1)));
    // !encrypted compound-OR: each condition needs to be the "first true" to cover its branch.
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                false,
                Optional.empty(),
                Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                false,
                Optional.empty(),
                Optional.empty(),
                Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                false,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(ExcelOoxmlChainingMode.CBC),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                false,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(256),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                false,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(16),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                false,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(100_000)));
    // encrypted path: null keyBits, null blockSize, null spinCount.
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                Optional.of(ExcelOoxmlChainingMode.CBC),
                Optional.empty(),
                Optional.of(16),
                Optional.of(1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                Optional.of(ExcelOoxmlChainingMode.CBC),
                Optional.of(256),
                Optional.empty(),
                Optional.of(1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlEncryptionSnapshot(
                true,
                Optional.of(ExcelOoxmlEncryptionMode.AGILE),
                Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                Optional.of(ExcelOoxmlChainingMode.CBC),
                Optional.of(256),
                Optional.of(16),
                Optional.empty()));

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
