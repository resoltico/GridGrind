package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.security.KeyStore;
import java.security.cert.Certificate;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Reflective and edge-path coverage for the OOXML package-security engine surface. */
class ExcelOoxmlPackageSecurityCoverageTest {
  @Test
  void materializeReadableWorkbookDistinguishesPlainEncryptedAndLegacyInputs() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-ooxml-materialize-");
    Path plainWorkbookPath = directory.resolve("plain.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Plain").setCell("A1", ExcelCellValue.text("Plain workbook"));
      workbook.save(plainWorkbookPath);
    }

    try (ExcelOoxmlPackageSecuritySupport.ReadableWorkbook readableWorkbook =
        ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
            plainWorkbookPath, null, Files::createTempFile)) {
      assertEquals(plainWorkbookPath.toAbsolutePath().normalize(), readableWorkbook.workbookPath());
      assertFalse(readableWorkbook.packageSecurity().isSecure());
      assertNull(readableWorkbook.sourceEncryptionPassword());
    }
    assertTrue(Files.exists(plainWorkbookPath));

    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(directory.resolve("encrypted"));
    Path[] decryptedPath = new Path[1];
    ExcelOoxmlPackageSecuritySupport.TempFileFactory tempFileFactory =
        (prefix, suffix) -> {
          decryptedPath[0] = Files.createTempFile(prefix, suffix);
          return decryptedPath[0];
        };

    try (ExcelOoxmlPackageSecuritySupport.ReadableWorkbook readableWorkbook =
        ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
            encryptedWorkbook.workbookPath(),
            new ExcelOoxmlOpenOptions(encryptedWorkbook.password()),
            tempFileFactory)) {
      assertNotEquals(
          encryptedWorkbook.workbookPath().toAbsolutePath().normalize(),
          readableWorkbook.workbookPath());
      assertEquals(encryptedWorkbook.password(), readableWorkbook.sourceEncryptionPassword());
      assertTrue(readableWorkbook.packageSecurity().encryption().encrypted());
      assertTrue(Files.exists(decryptedPath[0]));
    }
    assertFalse(Files.exists(decryptedPath[0]));

    Path legacyWorkbookPath = Files.createTempFile("gridgrind-legacy-materialize-", ".xls");
    try (HSSFWorkbook workbook = new HSSFWorkbook();
        OutputStream outputStream = Files.newOutputStream(legacyWorkbookPath)) {
      workbook.createSheet("Legacy").createRow(0).createCell(0).setCellValue("Legacy workbook");
      workbook.write(outputStream);
    }

    IllegalArgumentException legacyFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
                    legacyWorkbookPath,
                    new ExcelOoxmlOpenOptions("unused"),
                    Files::createTempFile));
    assertEquals("Only .xlsx workbooks are supported", legacyFailure.getMessage());
  }

  @Test
  void encryptedMaterializationCleansTempFilesAfterPasswordFailure() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-ooxml-materialize-cleanup-"));
    Path[] decryptedPath = new Path[1];

    InvalidWorkbookPasswordException failure =
        assertThrows(
            InvalidWorkbookPasswordException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
                    encryptedWorkbook.workbookPath(),
                    new ExcelOoxmlOpenOptions("wrong-password"),
                    (prefix, suffix) -> {
                      decryptedPath[0] = Files.createTempFile(prefix, suffix);
                      return decryptedPath[0];
                    }));

    assertEquals(encryptedWorkbook.workbookPath(), failure.workbookPath());
    assertNotNull(decryptedPath[0]);
    assertFalse(Files.exists(decryptedPath[0]));
  }

  @Test
  void encryptedMaterializationRequiresOpenOptionsPassword() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-ooxml-password-required-"));

    WorkbookPasswordRequiredException failure =
        assertThrows(
            WorkbookPasswordRequiredException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
                    encryptedWorkbook.workbookPath(), null, Files::createTempFile));

    assertEquals(encryptedWorkbook.workbookPath(), failure.workbookPath());

    WorkbookPasswordRequiredException nullPasswordFailure =
        assertThrows(
            WorkbookPasswordRequiredException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
                    encryptedWorkbook.workbookPath(),
                    new ExcelOoxmlOpenOptions(null),
                    Files::createTempFile));
    assertEquals(encryptedWorkbook.workbookPath(), nullPasswordFailure.workbookPath());
  }

  @Test
  void unsupportedMagicAndMaterializedWorkbookEdgeCasesStayExplicit() throws IOException {
    Path unsupportedPath = Files.createTempFile("gridgrind-unsupported-magic-", ".bin");
    Files.writeString(unsupportedPath, "plain text is not an OOXML workbook");
    IllegalArgumentException unsupportedMagicFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
                    unsupportedPath, null, Files::createTempFile));
    assertTrue(unsupportedMagicFailure.getMessage().contains("unsupported package magic"));

    Path missingWorkbookPath =
        Files.createTempDirectory("gridgrind-materialized-missing-").resolve("missing.xlsx");
    WorkbookNotFoundException missingPlainFailure =
        assertThrows(
            WorkbookNotFoundException.class,
            () ->
                ExcelWorkbook.openMaterializedWorkbook(
                    missingWorkbookPath,
                    missingWorkbookPath,
                    ExcelOoxmlPackageSecuritySnapshot.none(),
                    null));
    assertEquals(missingWorkbookPath.toAbsolutePath(), missingPlainFailure.workbookPath());

    WorkbookNotFoundException missingFormulaFailure =
        assertThrows(
            WorkbookNotFoundException.class,
            () ->
                ExcelWorkbook.openMaterializedWorkbook(
                    missingWorkbookPath,
                    ExcelFormulaEnvironment.defaults(),
                    missingWorkbookPath,
                    ExcelOoxmlPackageSecuritySnapshot.none(),
                    null));
    assertEquals(missingWorkbookPath.toAbsolutePath(), missingFormulaFailure.workbookPath());

    Path nonXlsxPath = Files.createTempFile("gridgrind-materialized-legacy-", ".xls");
    try (HSSFWorkbook workbook = new HSSFWorkbook();
        OutputStream outputStream = Files.newOutputStream(nonXlsxPath)) {
      workbook.createSheet("Legacy").createRow(0).createCell(0).setCellValue("Legacy workbook");
      workbook.write(outputStream);
    }

    IllegalArgumentException unsupportedMaterializedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelWorkbook.openMaterializedWorkbook(
                    nonXlsxPath, nonXlsxPath, ExcelOoxmlPackageSecuritySnapshot.none(), null));
    assertEquals("Only .xlsx workbooks are supported", unsupportedMaterializedFailure.getMessage());

    IllegalArgumentException unsupportedFormulaMaterializedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelWorkbook.openMaterializedWorkbook(
                    nonXlsxPath,
                    ExcelFormulaEnvironment.defaults(),
                    nonXlsxPath,
                    ExcelOoxmlPackageSecuritySnapshot.none(),
                    null));
    assertEquals(
        "Only .xlsx workbooks are supported", unsupportedFormulaMaterializedFailure.getMessage());
  }

  @Test
  void workbookOpenOverloadsCoverCustomTempFactoriesAndMaterializedPaths() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-workbook-open-security-"));
    AtomicInteger tempFilesCreated = new AtomicInteger();
    ExcelOoxmlPackageSecuritySupport.TempFileFactory tempFileFactory =
        (prefix, suffix) -> {
          tempFilesCreated.incrementAndGet();
          return Files.createTempFile(prefix, suffix);
        };

    try (ExcelWorkbook workbook =
            ExcelWorkbook.open(
                encryptedWorkbook.workbookPath(),
                new ExcelOoxmlOpenOptions(encryptedWorkbook.password()),
                tempFileFactory);
        ExcelWorkbook workbookWithEnvironment =
            ExcelWorkbook.open(
                encryptedWorkbook.workbookPath(),
                ExcelFormulaEnvironment.defaults(),
                new ExcelOoxmlOpenOptions(encryptedWorkbook.password()),
                tempFileFactory)) {
      assertEquals("Encrypted workbook", workbook.sheet("Encrypted").text("A1"));
      assertEquals("Encrypted workbook", workbookWithEnvironment.sheet("Encrypted").text("A1"));
    }
    assertEquals(2, tempFilesCreated.get());

    Path materializedWorkbookPath = Files.createTempFile("gridgrind-materialized-open-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Plain").setCell("A1", ExcelCellValue.text("Materialized"));
      workbook.save(materializedWorkbookPath);
    }

    try (ExcelWorkbook workbook =
            ExcelWorkbook.openMaterializedWorkbook(
                materializedWorkbookPath,
                materializedWorkbookPath,
                ExcelOoxmlPackageSecuritySnapshot.none(),
                null);
        ExcelWorkbook workbookWithEnvironment =
            ExcelWorkbook.openMaterializedWorkbook(
                materializedWorkbookPath,
                ExcelFormulaEnvironment.defaults(),
                materializedWorkbookPath,
                ExcelOoxmlPackageSecuritySnapshot.none(),
                null)) {
      assertEquals("Materialized", workbook.sheet("Plain").text("A1"));
      assertEquals("Materialized", workbookWithEnvironment.sheet("Plain").text("A1"));
    }
  }

  @Test
  void workbookOpenFailureHelpersCloseWorkbooksOnConstructionFailure() throws IOException {
    NullPointerException noFormulaFailure =
        assertThrows(
            NullPointerException.class,
            () -> ExcelWorkbook.openMaterializedWorkbook(new XSSFWorkbook(), null, null, null));
    assertEquals("loadedPackageSecurity must not be null", noFormulaFailure.getMessage());

    try (ThrowingOpenCloseWorkbook throwingNoFormulaWorkbook =
        new ThrowingOpenCloseWorkbook("no-formula close failure")) {
      try {
        NullPointerException throwingNoFormulaFailure =
            assertThrows(
                NullPointerException.class,
                () ->
                    ExcelWorkbook.openMaterializedWorkbook(
                        throwingNoFormulaWorkbook, null, null, null));
        assertEquals(1, throwingNoFormulaFailure.getSuppressed().length);
        assertEquals(
            "no-formula close failure", throwingNoFormulaFailure.getSuppressed()[0].getMessage());
      } finally {
        throwingNoFormulaWorkbook.disableCloseFailure();
      }
    }

    try (ThrowingOpenCloseWorkbook throwingFormulaWorkbook =
        new ThrowingOpenCloseWorkbook("formula close failure")) {
      try {
        NullPointerException formulaFailure =
            assertThrows(
                NullPointerException.class,
                () ->
                    ExcelWorkbook.openMaterializedWorkbook(
                        throwingFormulaWorkbook,
                        ExcelFormulaEnvironment.defaults(),
                        null,
                        null,
                        null));
        assertEquals("loadedPackageSecurity must not be null", formulaFailure.getMessage());
        assertEquals(1, formulaFailure.getSuppressed().length);
        assertEquals("formula close failure", formulaFailure.getSuppressed()[0].getMessage());
      } finally {
        throwingFormulaWorkbook.disableCloseFailure();
      }
    }

    RuntimeException runtimeFailure = new RuntimeException("close helper failure");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      ExcelWorkbook.closeWorkbookAfterOpenFailure(workbook, runtimeFailure);
      assertEquals(0, runtimeFailure.getSuppressed().length);
    }
  }

  @Test
  void securityHelpersCoverRemainingAliasAndFilesystemBranches() throws Exception {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            Files.createTempDirectory("gridgrind-signing-helper-branches-"));
    KeyStore signingKeyStore =
        loadPkcs12(signedWorkbook.pkcs12Path(), signedWorkbook.keystorePassword());
    Certificate certificate = signingKeyStore.getCertificate(signedWorkbook.alias());

    assertCopyDeleteAndEffectiveOptionsBranches(signedWorkbook.workbookPath());

    ExcelOoxmlSignatureOptions signatureOptions =
        new ExcelOoxmlSignatureOptions(
            signedWorkbook.pkcs12Path(),
            signedWorkbook.keystorePassword(),
            signedWorkbook.keyPassword(),
            signedWorkbook.alias(),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            null);
    assertSigningMaterialAndAliasBranches(signedWorkbook, signingKeyStore, signatureOptions);
    assertCertificateOnlyAndUninitializedKeystoreBranches(
        signedWorkbook, certificate, signatureOptions);
    assertSyntheticKeystoreBranches(signedWorkbook, signingKeyStore, certificate, signatureOptions);
  }

  @Test
  void failureTranslationHelpersCoverSecurityBridgeDefensivePaths() throws Exception {
    Path invalidEncryptedWorkbookPath =
        Files.createTempFile("gridgrind-invalid-encrypted-ooxml-", ".xlsx");
    try (POIFSFileSystem fileSystem = new POIFSFileSystem()) {
      fileSystem.createDocument(
          new ByteArrayInputStream(new byte[0]), Decryptor.DEFAULT_POIFS_ENTRY);
      try (OutputStream outputStream = Files.newOutputStream(invalidEncryptedWorkbookPath)) {
        fileSystem.writeFilesystem(outputStream);
      }
    }
    try (POIFSFileSystem invalidFileSystem =
        new POIFSFileSystem(invalidEncryptedWorkbookPath.toFile())) {
      IllegalArgumentException invalidEncryptionInfoFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelOoxmlPackageSecuritySupport.readEncryptionInfo(
                      invalidFileSystem, invalidEncryptedWorkbookPath));
      assertTrue(invalidEncryptionInfoFailure.getMessage().contains("not a supported encrypted"));
    }

    WorkbookSecurityException passwordFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.verifyPassword(
                    password -> {
                      throw new GeneralSecurityException("password failure");
                    },
                    "secret",
                    invalidEncryptedWorkbookPath));
    assertTrue(passwordFailure.getMessage().contains("verify the encrypted workbook password"));

    WorkbookSecurityException decryptFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.materializeDecryptedWorkbook(
                    () -> {
                      throw new GeneralSecurityException("decrypt failure");
                    },
                    Files.createTempFile("gridgrind-decrypt-failure-", ".xlsx"),
                    invalidEncryptedWorkbookPath));
    assertTrue(
        decryptFailure.getMessage().contains("Failed to decrypt the OOXML workbook package"));

    EncryptionInfo brokenEncryptionInfo = new EncryptionInfo(EncryptionMode.agile);
    brokenEncryptionInfo.setHeader(null);
    WorkbookSecurityException encryptionSnapshotFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () -> ExcelOoxmlPackageSecuritySupport.encryptionSnapshot(brokenEncryptionInfo));
    assertTrue(
        encryptionSnapshotFailure.getMessage().contains("inspect OOXML encryption metadata"));

    Path invalidPackagePath = Files.createTempFile("gridgrind-invalid-package-", ".txt");
    Files.writeString(invalidPackagePath, "not a package");
    WorkbookSecurityException packageInspectionFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.inspectPackageSecurity(
                    invalidPackagePath, ExcelOoxmlEncryptionSnapshot.none()));
    assertTrue(packageInspectionFailure.getMessage().contains("inspect OOXML package signatures"));

    assertNull(ExcelOoxmlPackageSecuritySupport.signerSubject(null));
    assertNull(ExcelOoxmlPackageSecuritySupport.signerIssuer(null));
    assertNull(ExcelOoxmlPackageSecuritySupport.signerSerialNumber(null));

    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            Files.createTempDirectory("gridgrind-signer-metadata-"));
    KeyStore signingKeyStore =
        loadPkcs12(signedWorkbook.pkcs12Path(), signedWorkbook.keystorePassword());
    java.security.cert.X509Certificate signer =
        (java.security.cert.X509Certificate) signingKeyStore.getCertificate(signedWorkbook.alias());
    assertTrue(
        ExcelOoxmlPackageSecuritySupport.signerSubject(signer).contains("GridGrind Signing Test"));
    assertTrue(
        ExcelOoxmlPackageSecuritySupport.signerIssuer(signer).contains("GridGrind Signing Test"));
    assertNotNull(ExcelOoxmlPackageSecuritySupport.signerSerialNumber(signer));

    WorkbookSecurityException signingFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.confirmAndVerifySignature(
                    () -> {
                      throw new javax.xml.crypto.dsig.XMLSignatureException("sign failure");
                    },
                    invalidEncryptedWorkbookPath));
    assertTrue(signingFailure.getMessage().contains("Failed to sign the OOXML workbook package"));

    WorkbookSecurityException runtimeSigningFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.confirmAndVerifySignature(
                    () -> {
                      throw new IllegalStateException("runtime sign failure");
                    },
                    invalidEncryptedWorkbookPath));
    assertTrue(runtimeSigningFailure.getMessage().contains("Unexpected OOXML signing failure"));

    WorkbookSecurityException invalidSignatureFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.confirmAndVerifySignature(
                    () -> false, invalidEncryptedWorkbookPath));
    assertTrue(invalidSignatureFailure.getMessage().contains("did not validate after signing"));

    WorkbookSecurityException encryptFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.writeEncryptedWorkbook(
                    fileSystem -> {
                      throw new GeneralSecurityException("encrypt failure");
                    },
                    Files.createTempFile("gridgrind-encrypt-source-", ".xlsx"),
                    Files.createTempFile("gridgrind-encrypt-target-", ".xlsx")));
    assertTrue(
        encryptFailure.getMessage().contains("Failed to encrypt the saved OOXML workbook package"));

    ExcelOoxmlSignatureOptions signatureOptions =
        new ExcelOoxmlSignatureOptions(
            signedWorkbook.pkcs12Path(),
            signedWorkbook.keystorePassword(),
            signedWorkbook.keyPassword(),
            null,
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            null);
    KeyStore singleAliasKeyStore = KeyStore.getInstance("PKCS12");
    singleAliasKeyStore.load(null, null);
    KeyStore.Entry keyEntry =
        signingKeyStore.getEntry(
            signedWorkbook.alias(),
            new KeyStore.PasswordProtection(signedWorkbook.keyPassword().toCharArray()));
    singleAliasKeyStore.setEntry(
        "only-alias",
        keyEntry,
        new KeyStore.PasswordProtection(signedWorkbook.keyPassword().toCharArray()));
    assertEquals(
        "only-alias",
        ExcelOoxmlPackageSecuritySupport.resolveAlias(singleAliasKeyStore, null, signatureOptions));

    Path malformedPackagePath = Files.createTempFile("gridgrind-sign-invalid-format-", ".xlsx");
    try (var outputStream =
        new java.util.zip.ZipOutputStream(Files.newOutputStream(malformedPackagePath))) {
      outputStream.putNextEntry(new java.util.zip.ZipEntry("broken.txt"));
      outputStream.write("broken".getBytes(java.nio.charset.StandardCharsets.UTF_8));
      outputStream.closeEntry();
    }
    WorkbookSecurityException invalidFormatFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.signWorkbook(
                    malformedPackagePath, signatureOptions));
    assertTrue(invalidFormatFailure.getMessage().contains("open the OOXML workbook package"));
  }

  @Test
  void saveAndPersistSecurityHelpersCoverPlainAndSignedGuardBranches() throws IOException {
    Path plainSourcePath = Files.createTempFile("gridgrind-saveworkbook-plain-source-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Plain").setCell("A1", ExcelCellValue.text("Plain save"));
      workbook.save(plainSourcePath);
    }

    Path plainSavedPath = plainSourcePath.resolveSibling("plain-saved-via-support.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.open(plainSourcePath)) {
      ExcelOoxmlPackageSecuritySupport.saveWorkbook(
          workbook,
          plainSavedPath,
          new ExcelOoxmlPersistenceOptions(null, null),
          Files::createTempFile);
    }
    try (ExcelWorkbook workbook = ExcelWorkbook.open(plainSavedPath)) {
      assertEquals("Plain save", workbook.sheet("Plain").text("A1"));
    }

    Path plainWorkbookPath = Files.createTempFile("gridgrind-persist-signed-plain-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Guard").setCell("A1", ExcelCellValue.text("Signed guard"));
      workbook.save(plainWorkbookPath);
    }

    ExcelOoxmlSignatureSnapshot signature =
        new ExcelOoxmlSignatureSnapshot(
            "/_xmlsignatures/sig1.xml",
            "CN=GridGrind Signing Test",
            "CN=GridGrind Signing Test",
            "01AB",
            ExcelOoxmlSignatureState.VALID);
    ExcelOoxmlPackageSecuritySnapshot signedSecurity =
        new ExcelOoxmlPackageSecuritySnapshot(
            ExcelOoxmlEncryptionSnapshot.none(), java.util.List.of(signature));

    IllegalArgumentException unsignedMutationFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.persistMaterializedWorkbook(
                    plainWorkbookPath,
                    plainWorkbookPath.resolveSibling("signed-guard-output.xlsx"),
                    signedSecurity,
                    null,
                    true,
                    new ExcelOoxmlPersistenceOptions(null, null)));
    assertTrue(unsignedMutationFailure.getMessage().contains("rewritten"));

    Path passThroughTarget = plainWorkbookPath.resolveSibling("signed-guard-unmutated-output.xlsx");
    ExcelOoxmlPackageSecuritySupport.persistMaterializedWorkbook(
        plainWorkbookPath,
        passThroughTarget,
        signedSecurity,
        null,
        false,
        new ExcelOoxmlPersistenceOptions(null, null));
    assertTrue(Files.exists(passThroughTarget));
  }

  @Test
  void saveWorkbookCoversInMemoryAndSignedPassthroughBranches() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Memory").setCell("A1", ExcelCellValue.text("Memory save"));
      Path memoryTarget = Files.createTempFile("gridgrind-save-memory-", ".xlsx");

      ExcelOoxmlPackageSecuritySupport.saveWorkbook(
          workbook, memoryTarget, null, Files::createTempFile);

      try (ExcelWorkbook reopened = ExcelWorkbook.open(memoryTarget)) {
        assertEquals("Memory save", reopened.sheet("Memory").text("A1"));
      }
    }

    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            Files.createTempDirectory("gridgrind-save-signed-passthrough-"));
    Path copiedWorkbookPath =
        signedWorkbook.workbookPath().getParent().resolve("copied-signed.xlsx");
    AtomicInteger tempFilesCreated = new AtomicInteger();
    try (ExcelWorkbook workbook = ExcelWorkbook.open(signedWorkbook.workbookPath())) {
      ExcelOoxmlPackageSecuritySupport.saveWorkbook(
          workbook,
          copiedWorkbookPath,
          new ExcelOoxmlPersistenceOptions(null, null),
          (prefix, suffix) -> {
            tempFilesCreated.incrementAndGet();
            return Files.createTempFile(prefix, suffix);
          });
    }
    assertEquals(0, tempFilesCreated.get());
    assertArrayEquals(
        Files.readAllBytes(signedWorkbook.workbookPath()), Files.readAllBytes(copiedWorkbookPath));
    assertTrue(OoxmlSecurityTestSupport.signatureValid(copiedWorkbookPath));

    Path resignedWorkbookPath = signedWorkbook.workbookPath().getParent().resolve("resigned.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.open(signedWorkbook.workbookPath())) {
      ExcelOoxmlPackageSecuritySupport.saveWorkbook(
          workbook,
          resignedWorkbookPath,
          new ExcelOoxmlPersistenceOptions(
              null,
              new ExcelOoxmlSignatureOptions(
                  signedWorkbook.pkcs12Path(),
                  signedWorkbook.keystorePassword(),
                  signedWorkbook.keyPassword(),
                  signedWorkbook.alias(),
                  ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                  null)),
          Files::createTempFile);
    }
    assertTrue(OoxmlSecurityTestSupport.signatureValid(resignedWorkbookPath));
  }

  private static void assertCopyDeleteAndEffectiveOptionsBranches(Path sourceWorkbookPath)
      throws IOException {
    ExcelOoxmlPackageSecuritySupport.copySourceWorkbook(sourceWorkbookPath, sourceWorkbookPath);
    assertTrue(OoxmlSecurityTestSupport.signatureValid(sourceWorkbookPath));

    Path copiedWorkbookPath = sourceWorkbookPath.getParent().resolve("copied-signed.xlsx");
    ExcelOoxmlPackageSecuritySupport.copySourceWorkbook(sourceWorkbookPath, copiedWorkbookPath);
    assertArrayEquals(
        Files.readAllBytes(sourceWorkbookPath), Files.readAllBytes(copiedWorkbookPath));

    Path deletedTempFile = Files.createTempFile("gridgrind-delete-if-exists-", ".tmp");
    ExcelOoxmlPackageSecuritySupport.deleteIfExists(null);
    ExcelOoxmlPackageSecuritySupport.deleteIfExists(deletedTempFile);
    assertFalse(Files.exists(deletedTempFile));

    Path nonEmptyDirectory = Files.createTempDirectory("gridgrind-delete-if-exists-dir-");
    Files.writeString(nonEmptyDirectory.resolve("child.txt"), "keep");
    ExcelOoxmlPackageSecuritySupport.deleteIfExists(nonEmptyDirectory);
    assertTrue(Files.exists(nonEmptyDirectory));

    ExcelOoxmlEncryptionSnapshot encryptedSnapshot =
        new ExcelOoxmlEncryptionSnapshot(
            true,
            ExcelOoxmlEncryptionMode.AGILE,
            "AES256",
            "SHA512",
            "ChainingModeCBC",
            256,
            16,
            100_000);
    IllegalStateException missingPasswordFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.effectiveOptions(
                    new ExcelOoxmlPackageSecuritySnapshot(encryptedSnapshot, java.util.List.of()),
                    null,
                    new ExcelOoxmlPersistenceOptions(null, null)));
    assertTrue(missingPasswordFailure.getMessage().contains("verified source password"));
  }

  private static void assertSigningMaterialAndAliasBranches(
      OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook,
      KeyStore signingKeyStore,
      ExcelOoxmlSignatureOptions signatureOptions)
      throws IOException, GeneralSecurityException {
    assertNotNull(ExcelOoxmlPackageSecuritySupport.signingMaterial(signatureOptions));

    ExcelOoxmlSignatureOptions wrongKeystorePassword =
        new ExcelOoxmlSignatureOptions(
            signedWorkbook.pkcs12Path(),
            "wrong-password",
            signedWorkbook.keyPassword(),
            signedWorkbook.alias(),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            null);
    InvalidSigningConfigurationException wrongKeystorePasswordFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.loadSigningKeyStore(
                    signedWorkbook.pkcs12Path(), wrongKeystorePassword));
    assertTrue(
        wrongKeystorePasswordFailure.getMessage().contains("Failed to load signing material"));

    InvalidSigningConfigurationException missingSigningMaterialFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.signingMaterial(
                    new ExcelOoxmlSignatureOptions(
                        signedWorkbook.pkcs12Path().resolveSibling("missing-signing-material.p12"),
                        signedWorkbook.keystorePassword(),
                        signedWorkbook.keyPassword(),
                        signedWorkbook.alias(),
                        ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                        null)));
    assertTrue(missingSigningMaterialFailure.getMessage().contains("does not exist"));

    assertEquals(
        signedWorkbook.alias(),
        ExcelOoxmlPackageSecuritySupport.resolveAlias(
            signingKeyStore, signedWorkbook.alias(), signatureOptions));

    InvalidSigningConfigurationException missingAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.resolveAlias(
                    signingKeyStore, "missing-alias", signatureOptions));
    assertTrue(missingAliasFailure.getMessage().contains("does not exist"));

    KeyStore emptyKeyStore = KeyStore.getInstance("PKCS12");
    emptyKeyStore.load(null, null);
    InvalidSigningConfigurationException noKeyAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.resolveAlias(
                    emptyKeyStore, null, signatureOptions));
    assertTrue(noKeyAliasFailure.getMessage().contains("does not contain a private-key entry"));
  }

  private static void assertCertificateOnlyAndUninitializedKeystoreBranches(
      OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook,
      Certificate certificate,
      ExcelOoxmlSignatureOptions signatureOptions)
      throws IOException, GeneralSecurityException {
    KeyStore certificateOnlyKeyStore = KeyStore.getInstance("PKCS12");
    certificateOnlyKeyStore.load(null, null);
    certificateOnlyKeyStore.setCertificateEntry("certificate-only", certificate);

    InvalidSigningConfigurationException nonPrivateKeyFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.signingPrivateKey(
                    certificateOnlyKeyStore,
                    signedWorkbook.pkcs12Path(),
                    "certificate-only",
                    signatureOptions));
    assertTrue(nonPrivateKeyFailure.getMessage().contains("does not resolve to a private key"));

    InvalidSigningConfigurationException certificateChainFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.signingCertificateChain(
                    certificateOnlyKeyStore, signedWorkbook.pkcs12Path(), "certificate-only"));
    assertTrue(
        certificateChainFailure
            .getMessage()
            .contains("does not contain an X.509 certificate chain"));

    KeyStore uninitializedKeyStore = KeyStore.getInstance("PKCS12");
    InvalidSigningConfigurationException aliasInspectionFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.resolveSigningAlias(
                    uninitializedKeyStore, signedWorkbook.pkcs12Path(), signatureOptions));
    assertTrue(aliasInspectionFailure.getMessage().contains("Failed to inspect signing aliases"));

    InvalidSigningConfigurationException privateKeyLoadFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.signingPrivateKey(
                    uninitializedKeyStore,
                    signedWorkbook.pkcs12Path(),
                    signedWorkbook.alias(),
                    signatureOptions));
    assertTrue(
        privateKeyLoadFailure.getMessage().contains("Failed to load the signing private key"));

    InvalidSigningConfigurationException chainLoadFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.signingCertificateChain(
                    uninitializedKeyStore, signedWorkbook.pkcs12Path(), signedWorkbook.alias()));
    assertTrue(
        chainLoadFailure.getMessage().contains("Failed to load the signing certificate chain"));
  }

  private static void assertSyntheticKeystoreBranches(
      OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook,
      KeyStore signingKeyStore,
      Certificate certificate,
      ExcelOoxmlSignatureOptions signatureOptions)
      throws IOException, GeneralSecurityException {
    KeyStore multiAliasKeyStore = KeyStore.getInstance("PKCS12");
    multiAliasKeyStore.load(null, null);
    KeyStore.Entry entry =
        signingKeyStore.getEntry(
            signedWorkbook.alias(),
            new KeyStore.PasswordProtection(signedWorkbook.keyPassword().toCharArray()));
    multiAliasKeyStore.setEntry(
        "first",
        entry,
        new KeyStore.PasswordProtection(signedWorkbook.keyPassword().toCharArray()));
    multiAliasKeyStore.setEntry(
        "second",
        entry,
        new KeyStore.PasswordProtection(signedWorkbook.keyPassword().toCharArray()));
    InvalidSigningConfigurationException multipleAliasesFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.resolveAlias(
                    multiAliasKeyStore, null, signatureOptions));
    assertTrue(multipleAliasesFailure.getMessage().contains("multiple private-key aliases"));

    KeyStore certificateOnlyAliasKeyStore = KeyStore.getInstance("PKCS12");
    certificateOnlyAliasKeyStore.load(null, null);
    certificateOnlyAliasKeyStore.setCertificateEntry("certificate-only", certificate);
    InvalidSigningConfigurationException certificateOnlyAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.resolveAlias(
                    certificateOnlyAliasKeyStore, null, signatureOptions));
    assertTrue(
        certificateOnlyAliasFailure.getMessage().contains("does not contain a private-key entry"));

    KeyStore nonPrivateKeyStore = fakeKeyStore(certificate.getPublicKey(), null);
    InvalidSigningConfigurationException nonPrivateAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.resolveAlias(
                    nonPrivateKeyStore, null, signatureOptions));
    assertTrue(
        nonPrivateAliasFailure.getMessage().contains("does not contain a private-key entry"));
    InvalidSigningConfigurationException reflectedNonPrivateKeyFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.signingPrivateKey(
                    nonPrivateKeyStore, signedWorkbook.pkcs12Path(), "alias", signatureOptions));
    assertTrue(reflectedNonPrivateKeyFailure.getMessage().contains("private key"));

    KeyStore nonX509ChainKeyStore =
        fakeKeyStore(
            signingKeyStore.getKey(
                signedWorkbook.alias(), signedWorkbook.keyPassword().toCharArray()),
            new Certificate[] {new DummyCertificate()});
    InvalidSigningConfigurationException nonX509Failure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.signingCertificateChain(
                    nonX509ChainKeyStore, signedWorkbook.pkcs12Path(), "alias"));
    assertTrue(nonX509Failure.getMessage().contains("non-X.509 certificate"));

    KeyStore emptyChainKeyStore =
        fakeKeyStore(
            signingKeyStore.getKey(
                signedWorkbook.alias(), signedWorkbook.keyPassword().toCharArray()),
            new Certificate[0]);
    InvalidSigningConfigurationException emptyChainFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.signingCertificateChain(
                    emptyChainKeyStore, signedWorkbook.pkcs12Path(), "alias"));
    assertTrue(emptyChainFailure.getMessage().contains("X.509 certificate chain"));

    Path signableWorkbookPath = Files.createTempFile("gridgrind-sign-description-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Signed").setCell("A1", ExcelCellValue.text("Signed workbook"));
      workbook.save(signableWorkbookPath);
    }
    ExcelOoxmlPackageSecuritySupport.signWorkbook(
        signableWorkbookPath,
        new ExcelOoxmlSignatureOptions(
            signedWorkbook.pkcs12Path(),
            signedWorkbook.keystorePassword(),
            signedWorkbook.keyPassword(),
            signedWorkbook.alias(),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            "GridGrind test signature"));
    assertTrue(OoxmlSecurityTestSupport.signatureValid(signableWorkbookPath));
  }

  private static KeyStore loadPkcs12(Path path, String password)
      throws IOException, GeneralSecurityException {
    KeyStore keyStore = KeyStore.getInstance("PKCS12");
    try (var inputStream = Files.newInputStream(path)) {
      keyStore.load(inputStream, password.toCharArray());
    }
    return keyStore;
  }

  private static KeyStore fakeKeyStore(java.security.Key key, Certificate[] certificateChain) {
    return new FakeKeyStore(new FakeKeyStoreSpi(key, certificateChain));
  }

  /** Workbook double that can fail `close()` while the open-failure helper is unwinding. */
  private static final class ThrowingOpenCloseWorkbook extends XSSFWorkbook {
    private final String closeMessage;
    private boolean failOnClose = true;

    private ThrowingOpenCloseWorkbook(String closeMessage) {
      this.closeMessage = closeMessage;
    }

    @Override
    public void close() throws IOException {
      if (failOnClose) {
        throw new IOException(closeMessage);
      }
      super.close();
    }

    private void disableCloseFailure() {
      failOnClose = false;
    }
  }

  /** Minimal initialized keystore wrapper for exercising alias and certificate edge cases. */
  private static final class FakeKeyStore extends KeyStore {
    private FakeKeyStore(FakeKeyStoreSpi keyStoreSpi) {
      super(keyStoreSpi, null, "fake");
      try {
        load(null, null);
      } catch (IOException | java.security.GeneralSecurityException exception) {
        throw new IllegalStateException("Failed to initialize the fake keystore", exception);
      }
    }
  }

  /** Minimal `KeyStoreSpi` implementation for deterministic non-private and non-X509 branches. */
  private static final class FakeKeyStoreSpi extends java.security.KeyStoreSpi {
    private final java.security.Key key;
    private final Certificate[] certificateChain;

    private FakeKeyStoreSpi(java.security.Key key, Certificate[] certificateChain) {
      this.key = key;
      this.certificateChain = certificateChain;
    }

    @Override
    public java.util.Enumeration<String> engineAliases() {
      return java.util.Collections.enumeration(java.util.List.of("alias"));
    }

    @Override
    public boolean engineContainsAlias(String alias) {
      return "alias".equals(alias);
    }

    @Override
    public int engineSize() {
      return 1;
    }

    @Override
    public boolean engineIsKeyEntry(String alias) {
      return key != null;
    }

    @Override
    public boolean engineIsCertificateEntry(String alias) {
      return key == null && certificateChain != null;
    }

    @Override
    public java.security.Key engineGetKey(String alias, char[] password)
        throws java.security.UnrecoverableKeyException {
      if (key == null) {
        return null;
      }
      return key;
    }

    @Override
    public Certificate[] engineGetCertificateChain(String alias) {
      return certificateChain == null ? null : certificateChain.clone();
    }

    @Override
    public Certificate engineGetCertificate(String alias) {
      return certificateChain == null || certificateChain.length == 0 ? null : certificateChain[0];
    }

    @Override
    public java.util.Date engineGetCreationDate(String alias) {
      return null;
    }

    @Override
    public void engineSetKeyEntry(
        String alias, java.security.Key key, char[] password, Certificate[] chain) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void engineSetKeyEntry(String alias, byte[] key, Certificate[] chain) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void engineSetCertificateEntry(String alias, Certificate cert) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void engineDeleteEntry(String alias) {
      throw new UnsupportedOperationException();
    }

    @Override
    public String engineGetCertificateAlias(Certificate cert) {
      return "alias";
    }

    @Override
    public void engineStore(OutputStream stream, char[] password) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void engineLoad(java.io.InputStream stream, char[] password) {}
  }

  /** Non-X509 certificate double for the signing-certificate-chain validation branch. */
  private static final class DummyCertificate extends Certificate {
    private static final long serialVersionUID = 1L;

    private DummyCertificate() {
      super("dummy");
    }

    @Override
    public byte[] getEncoded() {
      return new byte[0];
    }

    @Override
    public void verify(java.security.PublicKey key) {}

    @Override
    public void verify(java.security.PublicKey key, String sigProvider) {}

    @Override
    public String toString() {
      return "DummyCertificate";
    }

    @Override
    public java.security.PublicKey getPublicKey() {
      return null;
    }
  }
}
