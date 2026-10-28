package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureReport;
import dev.erst.gridgrind.core.model.security.ResolvedOpenSecurity;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.engine.security.SecurityAdapter;
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
        ResolvedOpenSecurity.Unencrypted.class, new ResolvedOpenSecurity.Unencrypted());
    assertEquals(EncryptionMode.agile, SecurityAdapter.toPoi(ExcelOoxmlEncryptionMode.AGILE));
    assertEquals(EncryptionMode.standard, SecurityAdapter.toPoi(ExcelOoxmlEncryptionMode.STANDARD));
    assertEquals(ExcelOoxmlEncryptionMode.AGILE, SecurityAdapter.fromPoi(EncryptionMode.agile));
    assertEquals(
        ExcelOoxmlEncryptionMode.STANDARD, SecurityAdapter.fromPoi(EncryptionMode.standard));
    assertThrows(
        IllegalArgumentException.class, () -> SecurityAdapter.fromPoi(EncryptionMode.binaryRC4));
    assertEquals(CipherAlgorithm.aes256, SecurityAdapter.toPoi(ExcelOoxmlWriteCipher.AES_256));
    assertEquals(CipherAlgorithm.aes192, SecurityAdapter.toPoi(ExcelOoxmlWriteCipher.AES_192));
    assertEquals(HashAlgorithm.sha512, SecurityAdapter.toPoi(ExcelOoxmlWriteHash.SHA_512));
    assertEquals(HashAlgorithm.sha384, SecurityAdapter.toPoi(ExcelOoxmlWriteHash.SHA_384));
    assertEquals(HashAlgorithm.sha256, SecurityAdapter.toPoi(ExcelOoxmlWriteHash.SHA_256));
    assertEquals(ChainingMode.ecb, SecurityAdapter.toPoi(ExcelOoxmlChainingMode.ECB));
    assertEquals(ChainingMode.cbc, SecurityAdapter.toPoi(ExcelOoxmlChainingMode.CBC));
    assertEquals(ChainingMode.cfb, SecurityAdapter.toPoi(ExcelOoxmlChainingMode.CFB));
    assertEquals(
        Optional.of(ExcelOoxmlWriteCipher.AES_256),
        SecurityAdapter.toWriteCipher(ExcelOoxmlCipherAlgorithm.AES_256));
    assertEquals(
        Optional.of(ExcelOoxmlWriteCipher.AES_192),
        SecurityAdapter.toWriteCipher(ExcelOoxmlCipherAlgorithm.AES_192));
    assertEquals(
        Optional.empty(), SecurityAdapter.toWriteCipher(ExcelOoxmlCipherAlgorithm.AES_128));
    assertEquals(
        Optional.of(ExcelOoxmlWriteHash.SHA_512),
        SecurityAdapter.toWriteHash(ExcelOoxmlHashAlgorithm.SHA_512));
    assertEquals(
        Optional.of(ExcelOoxmlWriteHash.SHA_384),
        SecurityAdapter.toWriteHash(ExcelOoxmlHashAlgorithm.SHA_384));
    assertEquals(
        Optional.of(ExcelOoxmlWriteHash.SHA_256),
        SecurityAdapter.toWriteHash(ExcelOoxmlHashAlgorithm.SHA_256));
    assertEquals(Optional.empty(), SecurityAdapter.toWriteHash(ExcelOoxmlHashAlgorithm.SHA_1));
    assertEquals(
        HashAlgorithm.sha256, SecurityAdapter.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA256));
    assertEquals(
        HashAlgorithm.sha384, SecurityAdapter.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA384));
    assertEquals(
        HashAlgorithm.sha512, SecurityAdapter.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA512));

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
    assertTrue(passwordRequired.getMessage().contains("requires a password"));
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
