package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureReport;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSecurityPoiBridge;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureOptions;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.apache.poi.poifs.crypt.ChainingMode;
import org.apache.poi.poifs.crypt.CipherAlgorithm;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.junit.jupiter.api.Test;

/** Direct validation tests for OOXML package-security engine value types. */
class ExcelOoxmlPackageSecurityTypesTest {
  @Test
  void openOptionsEncryptionOptionsAndSignatureOptionsNormalizeInputs() {
    String blankValue = Character.toString(' ');
    ExcelOoxmlOpenOptions openOptions = new ExcelOoxmlOpenOptions.Encrypted("secret");
    ExcelOoxmlEncryptionOptions encryptionOptions =
        new ExcelOoxmlEncryptionOptions("secret", null, null);
    ExcelOoxmlSignatureOptions signatureOptions =
        new ExcelOoxmlSignatureOptions(
            Path.of("/tmp/signing-material.p12"), "store-pass", null, null, null, null);

    assertEquals(
        "secret", assertInstanceOf(ExcelOoxmlOpenOptions.Encrypted.class, openOptions).password());
    assertEquals(ExcelOoxmlWriteCipher.AES_256, encryptionOptions.cipher());
    assertEquals(ExcelOoxmlWriteHash.SHA_512, encryptionOptions.hash());
    assertEquals("store-pass", signatureOptions.keyPassword());
    assertEquals(ExcelOoxmlSignatureDigestAlgorithm.SHA256, signatureOptions.digestAlgorithm());
    assertNull(signatureOptions.alias());
    assertNull(signatureOptions.description());

    assertInstanceOf(
        ExcelOoxmlOpenOptions.Unencrypted.class, new ExcelOoxmlOpenOptions.Unencrypted());
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelOoxmlOpenOptions.Encrypted(blankValue));
    assertThrows(NullPointerException.class, () -> new ExcelOoxmlOpenOptions.Encrypted(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelOoxmlEncryptionOptions(blankValue, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlSignatureOptions(
                Path.of("/tmp/signing-material.p12"), "store-pass", blankValue, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlSignatureOptions(
                Path.of("/tmp/signing-material.p12"), "store-pass", null, blankValue, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelOoxmlSignatureOptions(
                Path.of("/tmp/signing-material.p12"), "store-pass", null, null, null, blankValue));
  }

  @Test
  void encryptionAndSignatureSnapshotsValidateSecurityFacts() {
    String blankValue = Character.toString(' ');
    OoxmlEncryptionReport encryption =
        new OoxmlEncryptionReport.Encrypted(
            ExcelOoxmlEncryptionMode.AGILE,
            ExcelOoxmlCipherAlgorithm.AES_256,
            ExcelOoxmlHashAlgorithm.SHA_512,
            ExcelOoxmlChainingMode.CBC,
            256,
            16,
            100_000);
    OoxmlSignatureReport signature =
        new OoxmlSignatureReport(
            "/_xmlsignatures/sig1.xml",
            Optional.of("CN=GridGrind Signing Test"),
            Optional.of("CN=GridGrind Signing Test"),
            Optional.of("01AB"),
            ExcelOoxmlSignatureState.VALID);
    OoxmlPackageSecurityReport packageSecurity =
        new OoxmlPackageSecurityReport(encryption, List.of(signature));

    assertTrue(packageSecurity.isSecure());
    assertInstanceOf(
        OoxmlEncryptionReport.None.class, OoxmlPackageSecurityReport.none().encryption());
    assertFalse(OoxmlPackageSecurityReport.none().isSecure());
    assertEquals(
        Optional.empty(),
        new OoxmlSignatureReport(
                "/_xmlsignatures/sig2.xml",
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                ExcelOoxmlSignatureState.INVALID)
            .signer());
    assertEquals(
        Optional.of(
            new OoxmlSignatureReport.SignerIdentity(
                "CN=GridGrind Signing Test", "CN=GridGrind Signing Test", "01AB")),
        signature.signer());

    assertThrows(
        NullPointerException.class,
        () ->
            new OoxmlEncryptionReport.Encrypted(
                null,
                ExcelOoxmlCipherAlgorithm.AES_256,
                ExcelOoxmlHashAlgorithm.SHA_512,
                ExcelOoxmlChainingMode.CBC,
                256,
                16,
                1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlEncryptionReport.Encrypted(
                ExcelOoxmlEncryptionMode.AGILE,
                ExcelOoxmlCipherAlgorithm.AES_256,
                ExcelOoxmlHashAlgorithm.SHA_512,
                ExcelOoxmlChainingMode.CBC,
                0,
                16,
                1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlSignatureReport(
                blankValue,
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                ExcelOoxmlSignatureState.VALID));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlSignatureReport(
                "/_xmlsignatures/sig1.xml",
                Optional.of("CN=GridGrind Signing Test"),
                Optional.empty(),
                Optional.of("01AB"),
                ExcelOoxmlSignatureState.VALID));
    assertThrows(
        NullPointerException.class,
        () -> new OoxmlPackageSecurityReport(encryption, Arrays.asList(signature, null)));

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
        CipherAlgorithm.aes256, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlWriteCipher.AES_256));
    assertEquals(
        CipherAlgorithm.aes192, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlWriteCipher.AES_192));
    assertEquals(
        HashAlgorithm.sha512, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlWriteHash.SHA_512));
    assertEquals(
        HashAlgorithm.sha384, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlWriteHash.SHA_384));
    assertEquals(
        HashAlgorithm.sha256, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlWriteHash.SHA_256));
    assertEquals(ChainingMode.ecb, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlChainingMode.ECB));
    assertEquals(ChainingMode.cbc, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlChainingMode.CBC));
    assertEquals(ChainingMode.cfb, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlChainingMode.CFB));
    assertEquals(
        Optional.of(ExcelOoxmlWriteCipher.AES_256),
        ExcelOoxmlSecurityPoiBridge.toWriteCipher(ExcelOoxmlCipherAlgorithm.AES_256));
    assertEquals(
        Optional.of(ExcelOoxmlWriteCipher.AES_192),
        ExcelOoxmlSecurityPoiBridge.toWriteCipher(ExcelOoxmlCipherAlgorithm.AES_192));
    assertEquals(
        Optional.empty(),
        ExcelOoxmlSecurityPoiBridge.toWriteCipher(ExcelOoxmlCipherAlgorithm.AES_128));
    assertEquals(
        Optional.of(ExcelOoxmlWriteHash.SHA_512),
        ExcelOoxmlSecurityPoiBridge.toWriteHash(ExcelOoxmlHashAlgorithm.SHA_512));
    assertEquals(
        Optional.of(ExcelOoxmlWriteHash.SHA_384),
        ExcelOoxmlSecurityPoiBridge.toWriteHash(ExcelOoxmlHashAlgorithm.SHA_384));
    assertEquals(
        Optional.of(ExcelOoxmlWriteHash.SHA_256),
        ExcelOoxmlSecurityPoiBridge.toWriteHash(ExcelOoxmlHashAlgorithm.SHA_256));
    assertEquals(
        Optional.empty(), ExcelOoxmlSecurityPoiBridge.toWriteHash(ExcelOoxmlHashAlgorithm.SHA_1));
    assertEquals(
        HashAlgorithm.sha256,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA256));
    assertEquals(
        HashAlgorithm.sha384,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA384));
    assertEquals(
        HashAlgorithm.sha512,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA512));

    OoxmlEncryptionReport none = new OoxmlEncryptionReport.None();
    assertInstanceOf(OoxmlEncryptionReport.None.class, none);
    assertThrows(
        NullPointerException.class,
        () ->
            new OoxmlEncryptionReport.Encrypted(
                ExcelOoxmlEncryptionMode.AGILE,
                null,
                ExcelOoxmlHashAlgorithm.SHA_512,
                ExcelOoxmlChainingMode.CBC,
                256,
                16,
                1));
    assertThrows(
        NullPointerException.class,
        () ->
            new OoxmlEncryptionReport.Encrypted(
                ExcelOoxmlEncryptionMode.AGILE,
                ExcelOoxmlCipherAlgorithm.AES_256,
                null,
                ExcelOoxmlChainingMode.CBC,
                256,
                16,
                1));
    assertThrows(
        NullPointerException.class,
        () ->
            new OoxmlEncryptionReport.Encrypted(
                ExcelOoxmlEncryptionMode.AGILE,
                ExcelOoxmlCipherAlgorithm.AES_256,
                ExcelOoxmlHashAlgorithm.SHA_512,
                null,
                256,
                16,
                1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlEncryptionReport.Encrypted(
                ExcelOoxmlEncryptionMode.AGILE,
                ExcelOoxmlCipherAlgorithm.AES_256,
                ExcelOoxmlHashAlgorithm.SHA_512,
                ExcelOoxmlChainingMode.CBC,
                256,
                0,
                1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new OoxmlEncryptionReport.Encrypted(
                ExcelOoxmlEncryptionMode.AGILE,
                ExcelOoxmlCipherAlgorithm.AES_256,
                ExcelOoxmlHashAlgorithm.SHA_512,
                ExcelOoxmlChainingMode.CBC,
                256,
                16,
                -1));

    Path workbookPath = Path.of("/tmp/security.xlsx");
    WorkbookPasswordRequiredException passwordRequired = new WorkbookPasswordRequiredException();
    assertTrue(passwordRequired.getMessage().contains("source.security.password"));
    assertFalse(passwordRequired.getMessage().contains(workbookPath.toString()));

    InvalidWorkbookPasswordException invalidPassword = new InvalidWorkbookPasswordException();
    assertTrue(invalidPassword.getMessage().contains("did not unlock the source workbook"));
    assertFalse(invalidPassword.getMessage().contains(workbookPath.toString()));

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
