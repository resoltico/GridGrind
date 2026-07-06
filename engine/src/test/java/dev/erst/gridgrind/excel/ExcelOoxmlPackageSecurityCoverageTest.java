package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionSnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageEncryptionSupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageFileSupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageInspectionSupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackagePersistenceSupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSigningSupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureSnapshot;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.security.KeyStore;
import java.security.cert.Certificate;
import java.util.Optional;
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
    Path directory = ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-materialize-");
    Path plainWorkbookPath = directory.resolve("plain.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Plain")
          .cells()
          .setCell("A1", ExcelCellValue.text("Plain workbook"));
      workbook
          .persistence()
          .save(
              plainWorkbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelOoxmlPackageSecuritySupport.ReadableWorkbook readableWorkbook =
        ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
            plainWorkbookPath, null, Files::createTempFile)) {
      assertEquals(plainWorkbookPath.toAbsolutePath().normalize(), readableWorkbook.workbookPath());
      assertFalse(readableWorkbook.packageSecurity().isSecure());
      assertEquals(Optional.empty(), readableWorkbook.sourceEncryptionPassword());
    }
    assertTrue(Files.exists(plainWorkbookPath));

    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(directory.resolve("encrypted"));
    Path explicitTempRoot = directory.resolve("caller-scratch");
    AtomicInteger tempFilesCreated = new AtomicInteger();
    WorkbookTempFileFactory tempFileFactory =
        (prefix, suffix) -> {
          tempFilesCreated.incrementAndGet();
          return ExcelTempFiles.createManagedTempFile(explicitTempRoot, prefix, suffix);
        };

    try (ExcelOoxmlPackageSecuritySupport.ReadableWorkbook readableWorkbook =
        ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
            encryptedWorkbook.workbookPath(),
            new ExcelOoxmlOpenOptions.Encrypted(encryptedWorkbook.password()),
            tempFileFactory)) {
      assertNotEquals(
          encryptedWorkbook.workbookPath().toAbsolutePath().normalize(),
          readableWorkbook.workbookPath());
      assertEquals(
          Optional.of(encryptedWorkbook.password()), readableWorkbook.sourceEncryptionPassword());
      assertInstanceOf(
          ExcelOoxmlEncryptionSnapshot.Encrypted.class,
          readableWorkbook.packageSecurity().encryption());
      assertTrue(Files.exists(readableWorkbook.workbookPath()));
      assertFalse(
          readableWorkbook
              .workbookPath()
              .startsWith(explicitTempRoot.toAbsolutePath().normalize()));
    }
    assertEquals(0, tempFilesCreated.get());
    assertTrue(Files.notExists(explicitTempRoot));

    Path legacyWorkbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-legacy-materialize-", ".xls");
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
                    new ExcelOoxmlOpenOptions.Encrypted("unused"),
                    Files::createTempFile));
    assertEquals("Only .xlsx workbooks are supported", legacyFailure.getMessage());
  }

  @Test
  void workbookPackageIoBridgesMaterializationAndPersistenceForPlainWorkbooks() throws IOException {
    Path directory = ExcelTempFiles.createManagedTempDirectory("gridgrind-package-io-");
    Path sourceWorkbookPath = directory.resolve("source.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", ExcelCellValue.text("Bridge"));
      workbook
          .persistence()
          .save(
              sourceWorkbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    Path copiedWorkbookPath = directory.resolve("copied.xlsx");
    try (WorkbookPackageIo.MaterializedWorkbook materializedWorkbook =
        WorkbookPackageIo.materializeWorkbook(sourceWorkbookPath, null, Files::createTempFile)) {
      assertEquals(
          sourceWorkbookPath.toAbsolutePath().normalize(), materializedWorkbook.workbookPath());
      WorkbookPackageIo.persistMaterializedWorkbook(
          materializedWorkbook.workbookPath(),
          copiedWorkbookPath,
          ExcelOoxmlPackageSecuritySnapshot.none(),
          Optional.empty(),
          false,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          ExcelOoxmlPersistenceOptions.none());
    }

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(
            copiedWorkbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Bridge", reopened.sheet("Budget").cells().text("A1"));
    }
  }

  @Test
  void workbookFileSupportCoversReplaceExistingCopyAndMoveBranches() throws IOException {
    Path workspace = ExcelTempFiles.createManagedTempDirectory("gridgrind-package-file-ops-");
    Path sourceWorkbookPath = workspace.resolve("source.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", ExcelCellValue.text("Source"));
      workbook
          .persistence()
          .save(
              sourceWorkbookPath,
              WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    Path replaceExistingTarget = workspace.resolve("replace-existing.xlsx");
    Files.writeString(replaceExistingTarget, "old");
    ExcelOoxmlPackageFileSupport.copySourceWorkbook(
        sourceWorkbookPath,
        replaceExistingTarget,
        WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
    assertArrayEquals(
        Files.readAllBytes(sourceWorkbookPath), Files.readAllBytes(replaceExistingTarget));
    NullPointerException nullTargetFailure =
        assertThrows(
            NullPointerException.class,
            () ->
                ExcelOoxmlPackageFileSupport.copySourceWorkbook(
                    sourceWorkbookPath, null, WorkbookArtifactWriteDisposition.CREATE_NEW));
    assertEquals("targetPath must not be null", nullTargetFailure.getMessage());
    ExcelOoxmlPackageFileSupport.copySourceWorkbook(
        sourceWorkbookPath, sourceWorkbookPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
    assertTrue(Files.exists(sourceWorkbookPath));

    Path moveCreateNewSource = workspace.resolve("move-create-new-source.xlsx");
    Files.copy(sourceWorkbookPath, moveCreateNewSource);
    Path moveCreateNewTarget = workspace.resolve("move-create-new-target.xlsx");
    ExcelOoxmlPackageFileSupport.moveWorkbook(
        moveCreateNewSource, moveCreateNewTarget, WorkbookArtifactWriteDisposition.CREATE_NEW);
    assertFalse(Files.exists(moveCreateNewSource));
    assertTrue(Files.exists(moveCreateNewTarget));

    Path moveReplaceSource = workspace.resolve("move-replace-source.xlsx");
    Files.copy(sourceWorkbookPath, moveReplaceSource);
    Path moveReplaceTarget = workspace.resolve("move-replace-target.xlsx");
    Files.writeString(moveReplaceTarget, "old target");
    ExcelOoxmlPackageFileSupport.moveWorkbook(
        moveReplaceSource, moveReplaceTarget, WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
    assertFalse(Files.exists(moveReplaceSource));
    assertArrayEquals(
        Files.readAllBytes(sourceWorkbookPath), Files.readAllBytes(moveReplaceTarget));

    Path noOpMovePath = workspace.resolve("move-no-op.xlsx");
    Files.copy(sourceWorkbookPath, noOpMovePath);
    ExcelOoxmlPackageFileSupport.moveWorkbook(
        noOpMovePath, noOpMovePath, WorkbookArtifactWriteDisposition.CREATE_NEW);
    assertTrue(Files.exists(noOpMovePath));
  }

  @Test
  void encryptedMaterializationCleansTempFilesAfterPasswordFailure() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-materialize-cleanup-"));
    Path explicitTempRoot = encryptedWorkbook.workbookPath().getParent().resolve("caller-scratch");
    AtomicInteger tempFilesCreated = new AtomicInteger();

    InvalidWorkbookPasswordException failure =
        assertThrows(
            InvalidWorkbookPasswordException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
                    encryptedWorkbook.workbookPath(),
                    new ExcelOoxmlOpenOptions.Encrypted("wrong-password"),
                    (prefix, suffix) -> {
                      tempFilesCreated.incrementAndGet();
                      return ExcelTempFiles.createManagedTempFile(explicitTempRoot, prefix, suffix);
                    }));

    assertEquals(encryptedWorkbook.workbookPath(), failure.workbookPath());
    assertEquals(0, tempFilesCreated.get());
    assertTrue(Files.notExists(explicitTempRoot));
  }

  @Test
  void encryptedMaterializationRequiresOpenOptionsPassword() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-password-required-"));

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
                    new ExcelOoxmlOpenOptions.Unencrypted(),
                    Files::createTempFile));
    assertEquals(encryptedWorkbook.workbookPath(), nullPasswordFailure.workbookPath());
  }

  @Test
  void unsupportedMagicAndMaterializedWorkbookEdgeCasesStayExplicit() throws IOException {
    Path unsupportedPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-unsupported-magic-", ".bin");
    Files.writeString(unsupportedPath, "plain text is not an OOXML workbook");
    IllegalArgumentException unsupportedMagicFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
                    unsupportedPath, null, Files::createTempFile));
    assertTrue(unsupportedMagicFailure.getMessage().contains("unsupported package magic"));

    Path missingWorkbookPath =
        ExcelTempFiles.createManagedTempDirectory("gridgrind-materialized-missing-")
            .resolve("missing.xlsx");
    WorkbookNotFoundException missingPlainFailure =
        assertThrows(
            WorkbookNotFoundException.class,
            () ->
                ExcelWorkbooks.openMaterializedWorkbook(
                    missingWorkbookPath,
                    Optional.of(missingWorkbookPath),
                    ExcelOoxmlPackageSecuritySnapshot.none(),
                    Optional.empty()));
    assertEquals(missingWorkbookPath.toAbsolutePath(), missingPlainFailure.workbookPath());

    WorkbookNotFoundException missingFormulaFailure =
        assertThrows(
            WorkbookNotFoundException.class,
            () ->
                ExcelWorkbooks.openMaterializedWorkbook(
                    missingWorkbookPath,
                    ExcelFormulaEnvironment.defaults(),
                    Optional.of(missingWorkbookPath),
                    ExcelOoxmlPackageSecuritySnapshot.none(),
                    Optional.empty()));
    assertEquals(missingWorkbookPath.toAbsolutePath(), missingFormulaFailure.workbookPath());

    Path nonXlsxPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-materialized-legacy-", ".xls");
    try (HSSFWorkbook workbook = new HSSFWorkbook();
        OutputStream outputStream = Files.newOutputStream(nonXlsxPath)) {
      workbook.createSheet("Legacy").createRow(0).createCell(0).setCellValue("Legacy workbook");
      workbook.write(outputStream);
    }

    IllegalArgumentException unsupportedMaterializedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelWorkbooks.openMaterializedWorkbook(
                    nonXlsxPath,
                    Optional.of(nonXlsxPath),
                    ExcelOoxmlPackageSecuritySnapshot.none(),
                    Optional.empty()));
    assertEquals("Only .xlsx workbooks are supported", unsupportedMaterializedFailure.getMessage());

    IllegalArgumentException unsupportedFormulaMaterializedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelWorkbooks.openMaterializedWorkbook(
                    nonXlsxPath,
                    ExcelFormulaEnvironment.defaults(),
                    Optional.of(nonXlsxPath),
                    ExcelOoxmlPackageSecuritySnapshot.none(),
                    Optional.empty()));
    assertEquals(
        "Only .xlsx workbooks are supported", unsupportedFormulaMaterializedFailure.getMessage());
  }

  @Test
  void workbookOpenOverloadsCoverCustomTempFactoriesAndMaterializedPaths() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-workbook-open-security-"));
    AtomicInteger tempFilesCreated = new AtomicInteger();
    WorkbookTempFileFactory tempFileFactory =
        (prefix, suffix) -> {
          tempFilesCreated.incrementAndGet();
          return ExcelTempFiles.createManagedTempFile(prefix, suffix);
        };

    try (ExcelWorkbook workbook =
            ExcelWorkbooks.open(
                encryptedWorkbook.workbookPath(),
                new ExcelOoxmlOpenOptions.Encrypted(encryptedWorkbook.password()),
                tempFileFactory);
        ExcelWorkbook workbookWithEnvironment =
            ExcelWorkbooks.open(
                encryptedWorkbook.workbookPath(),
                ExcelFormulaEnvironment.defaults(),
                new ExcelOoxmlOpenOptions.Encrypted(encryptedWorkbook.password()),
                tempFileFactory)) {
      assertEquals("Encrypted workbook", workbook.sheet("Encrypted").cells().text("A1"));
      assertEquals(
          "Encrypted workbook", workbookWithEnvironment.sheet("Encrypted").cells().text("A1"));
    }
    assertEquals(0, tempFilesCreated.get());

    Path materializedWorkbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-materialized-open-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Plain").cells().setCell("A1", ExcelCellValue.text("Materialized"));
      workbook
          .persistence()
          .save(
              materializedWorkbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (ExcelWorkbook workbook =
            ExcelWorkbooks.openMaterializedWorkbook(
                materializedWorkbookPath,
                Optional.of(materializedWorkbookPath),
                ExcelOoxmlPackageSecuritySnapshot.none(),
                Optional.empty());
        ExcelWorkbook workbookWithEnvironment =
            ExcelWorkbooks.openMaterializedWorkbook(
                materializedWorkbookPath,
                ExcelFormulaEnvironment.defaults(),
                Optional.of(materializedWorkbookPath),
                ExcelOoxmlPackageSecuritySnapshot.none(),
                Optional.empty())) {
      assertEquals("Materialized", workbook.sheet("Plain").cells().text("A1"));
      assertEquals("Materialized", workbookWithEnvironment.sheet("Plain").cells().text("A1"));
    }
  }

  @Test
  void workbookOpenFailureHelpersCloseWorkbooksOnConstructionFailure() throws IOException {
    NullPointerException noFormulaFailure =
        assertThrows(
            NullPointerException.class,
            () ->
                ExcelWorkbookOpenSupport.openMaterializedWorkbook(
                    new XSSFWorkbook(), Optional.empty(), null, Optional.empty()));
    assertEquals("loadedPackageSecurity must not be null", noFormulaFailure.getMessage());

    try (ThrowingOpenCloseWorkbook throwingNoFormulaWorkbook =
        new ThrowingOpenCloseWorkbook("no-formula close failure")) {
      try {
        NullPointerException throwingNoFormulaFailure =
            assertThrows(
                NullPointerException.class,
                () ->
                    ExcelWorkbookOpenSupport.openMaterializedWorkbook(
                        throwingNoFormulaWorkbook, Optional.empty(), null, Optional.empty()));
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
                    ExcelWorkbookOpenSupport.openMaterializedWorkbook(
                        throwingFormulaWorkbook,
                        ExcelFormulaEnvironment.defaults(),
                        Optional.empty(),
                        null,
                        Optional.empty()));
        assertEquals("loadedPackageSecurity must not be null", formulaFailure.getMessage());
        assertEquals(1, formulaFailure.getSuppressed().length);
        assertEquals("formula close failure", formulaFailure.getSuppressed()[0].getMessage());
      } finally {
        throwingFormulaWorkbook.disableCloseFailure();
      }
    }

    RuntimeException runtimeFailure = new RuntimeException("close helper failure");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      ExcelWorkbooks.closeWorkbookAfterOpenFailure(workbook, runtimeFailure);
      assertEquals(0, runtimeFailure.getSuppressed().length);
    }
  }

  @Test
  void securityHelpersCoverRemainingAliasAndFilesystemBranches() throws Exception {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-signing-helper-branches-"));
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
        ExcelTempFiles.createManagedTempFile("gridgrind-invalid-encrypted-ooxml-", ".xlsx");
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
                  ExcelOoxmlPackageEncryptionSupport.readEncryptionInfo(
                      invalidFileSystem, invalidEncryptedWorkbookPath));
      assertTrue(invalidEncryptionInfoFailure.getMessage().contains("not a supported encrypted"));
    }

    WorkbookSecurityException passwordFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageEncryptionSupport.verifyPassword(
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
                ExcelOoxmlPackageEncryptionSupport.materializeDecryptedWorkbook(
                    () -> {
                      throw new GeneralSecurityException("decrypt failure");
                    },
                    ExcelTempFiles.createManagedTempFile("gridgrind-decrypt-failure-", ".xlsx"),
                    invalidEncryptedWorkbookPath));
    assertTrue(
        decryptFailure.getMessage().contains("Failed to decrypt the OOXML workbook package"));

    EncryptionInfo brokenEncryptionInfo = new EncryptionInfo(EncryptionMode.agile);
    brokenEncryptionInfo.setHeader(null);
    WorkbookSecurityException encryptionSnapshotFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () -> ExcelOoxmlPackageEncryptionSupport.encryptionSnapshot(brokenEncryptionInfo));
    assertTrue(
        encryptionSnapshotFailure.getMessage().contains("inspect OOXML encryption metadata"));

    Path invalidPackagePath =
        ExcelTempFiles.createManagedTempFile("gridgrind-invalid-package-", ".txt");
    Files.writeString(invalidPackagePath, "not a package");
    WorkbookSecurityException packageInspectionFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageInspectionSupport.inspectPackageSecurity(
                    invalidPackagePath, ExcelOoxmlEncryptionSnapshot.none()));
    assertTrue(packageInspectionFailure.getMessage().contains("inspect OOXML package signatures"));

    assertEquals(Optional.empty(), ExcelOoxmlPackageInspectionSupport.signerIdentity(null));

    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-signer-metadata-"));
    KeyStore signingKeyStore =
        loadPkcs12(signedWorkbook.pkcs12Path(), signedWorkbook.keystorePassword());
    java.security.cert.X509Certificate signer =
        (java.security.cert.X509Certificate) signingKeyStore.getCertificate(signedWorkbook.alias());
    ExcelOoxmlSignatureSnapshot.SignerIdentity signerIdentity =
        ExcelOoxmlPackageInspectionSupport.signerIdentity(signer).orElseThrow();
    assertTrue(signerIdentity.subject().contains("GridGrind Signing Test"));
    assertTrue(signerIdentity.issuer().contains("GridGrind Signing Test"));
    assertFalse(signerIdentity.serialNumberHex().isBlank());

    ExcelOoxmlSignatureSnapshot inspectedSignature =
        ExcelOoxmlPackageInspectionSupport.inspectPackageSecurity(
                signedWorkbook.workbookPath(), ExcelOoxmlEncryptionSnapshot.none())
            .signatures()
            .getFirst();
    assertEquals(ExcelOoxmlSignatureState.VALID, inspectedSignature.state());
    assertTrue(
        inspectedSignature.signer().orElseThrow().subject().contains("GridGrind Signing Test"));

    WorkbookSecurityException signingFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.confirmAndVerifySignature(
                    () -> {
                      throw new javax.xml.crypto.dsig.XMLSignatureException("sign failure");
                    },
                    invalidEncryptedWorkbookPath));
    assertTrue(signingFailure.getMessage().contains("Failed to sign the OOXML workbook package"));

    WorkbookSecurityException runtimeSigningFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.confirmAndVerifySignature(
                    () -> {
                      throw new IllegalStateException("runtime sign failure");
                    },
                    invalidEncryptedWorkbookPath));
    assertTrue(runtimeSigningFailure.getMessage().contains("Unexpected OOXML signing failure"));

    WorkbookSecurityException invalidSignatureFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.confirmAndVerifySignature(
                    () -> false, invalidEncryptedWorkbookPath));
    assertTrue(invalidSignatureFailure.getMessage().contains("did not validate after signing"));

    WorkbookSecurityException encryptFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                ExcelOoxmlPackageEncryptionSupport.writeEncryptedWorkbook(
                    fileSystem -> {
                      throw new GeneralSecurityException("encrypt failure");
                    },
                    ExcelTempFiles.createManagedTempFile("gridgrind-encrypt-source-", ".xlsx"),
                    ExcelTempFiles.createManagedTempFile("gridgrind-encrypt-target-", ".xlsx"),
                    WorkbookArtifactWriteDisposition.CREATE_NEW));
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
        ExcelOoxmlPackageSigningSupport.resolveAlias(singleAliasKeyStore, null, signatureOptions));

    Path malformedPackagePath =
        ExcelTempFiles.createManagedTempFile("gridgrind-sign-invalid-format-", ".xlsx");
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
                ExcelOoxmlPackageSigningSupport.signWorkbook(
                    malformedPackagePath, signatureOptions));
    assertTrue(invalidFormatFailure.getMessage().contains("open the OOXML workbook package"));
  }

  @Test
  void saveAndPersistSecurityHelpersCoverPlainAndSignedGuardBranches() throws IOException {
    Path workspace = ExcelTempFiles.createManagedTempDirectory("gridgrind-saveworkbook-security-");
    Path plainSourcePath = workspace.resolve("plain-source.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Plain").cells().setCell("A1", ExcelCellValue.text("Plain save"));
      workbook
          .persistence()
          .save(
              plainSourcePath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    Path plainSavedPath = workspace.resolve("plain-saved-via-support.xlsx");
    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(plainSourcePath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelOoxmlPackageSecuritySupport.saveWorkbook(
          workbook,
          plainSavedPath,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          ExcelOoxmlPersistenceOptions.none(),
          Files::createTempFile);
    }
    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(plainSavedPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Plain save", workbook.sheet("Plain").cells().text("A1"));
    }

    Path plainWorkbookPath = workspace.resolve("persist-signed-plain.xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Guard").cells().setCell("A1", ExcelCellValue.text("Signed guard"));
      workbook
          .persistence()
          .save(
              plainWorkbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    ExcelOoxmlSignatureSnapshot signature =
        new ExcelOoxmlSignatureSnapshot(
            "/_xmlsignatures/sig1.xml",
            Optional.of("CN=GridGrind Signing Test"),
            Optional.of("CN=GridGrind Signing Test"),
            Optional.of("01AB"),
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
                    Optional.empty(),
                    true,
                    WorkbookArtifactWriteDisposition.CREATE_NEW,
                    ExcelOoxmlPersistenceOptions.none()));
    assertTrue(unsignedMutationFailure.getMessage().contains("rewritten"));

    Path passThroughTarget = plainWorkbookPath.resolveSibling("signed-guard-unmutated-output.xlsx");
    ExcelOoxmlPackageSecuritySupport.persistMaterializedWorkbook(
        plainWorkbookPath,
        passThroughTarget,
        signedSecurity,
        Optional.empty(),
        false,
        WorkbookArtifactWriteDisposition.CREATE_NEW,
        ExcelOoxmlPersistenceOptions.none());
    assertTrue(Files.exists(passThroughTarget));
  }

  @Test
  void saveWorkbookCoversInMemoryAndSignedPassthroughBranches() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Memory").cells().setCell("A1", ExcelCellValue.text("Memory save"));
      Path memoryTarget = ExcelTempFiles.createManagedTempFile("gridgrind-save-memory-", ".xlsx");

      ExcelOoxmlPackageSecuritySupport.saveWorkbook(
          workbook,
          memoryTarget,
          WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
          null,
          Files::createTempFile);

      try (ExcelWorkbook reopened =
          ExcelWorkbooks.open(memoryTarget, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
        assertEquals("Memory save", reopened.sheet("Memory").cells().text("A1"));
      }
    }

    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-save-signed-passthrough-"));
    Path copiedWorkbookPath =
        signedWorkbook.workbookPath().getParent().resolve("copied-signed.xlsx");
    AtomicInteger tempFilesCreated = new AtomicInteger();
    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            signedWorkbook.workbookPath(), ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelOoxmlPackageSecuritySupport.saveWorkbook(
          workbook,
          copiedWorkbookPath,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          ExcelOoxmlPersistenceOptions.none(),
          (prefix, suffix) -> {
            tempFilesCreated.incrementAndGet();
            return ExcelTempFiles.createManagedTempFile(prefix, suffix);
          });
    }
    assertEquals(0, tempFilesCreated.get());
    assertArrayEquals(
        Files.readAllBytes(signedWorkbook.workbookPath()), Files.readAllBytes(copiedWorkbookPath));
    assertTrue(OoxmlSecurityTestSupport.signatureValid(copiedWorkbookPath));

    Path resignedWorkbookPath = signedWorkbook.workbookPath().getParent().resolve("resigned.xlsx");
    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            signedWorkbook.workbookPath(), ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      ExcelOoxmlPackageSecuritySupport.saveWorkbook(
          workbook,
          resignedWorkbookPath,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          new ExcelOoxmlPersistenceOptions(
              Optional.empty(),
              Optional.of(
                  new ExcelOoxmlSignatureOptions(
                      signedWorkbook.pkcs12Path(),
                      signedWorkbook.keystorePassword(),
                      signedWorkbook.keyPassword(),
                      signedWorkbook.alias(),
                      ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                      null))),
          Files::createTempFile);
    }
    assertTrue(OoxmlSecurityTestSupport.signatureValid(resignedWorkbookPath));
  }

  @Test
  void encryptedSaveKeepsPlaintextTempsOutOfCallerScratchRoot() throws IOException {
    Path workspace = ExcelTempFiles.createManagedTempDirectory("gridgrind-encrypted-save-private-");
    Path explicitTempRoot = workspace.resolve("caller-scratch");
    Path encryptedOutput = workspace.resolve("encrypted-output.xlsx");
    AtomicInteger tempFilesCreated = new AtomicInteger();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Encrypted")
          .cells()
          .setCell("A1", ExcelCellValue.text("Encrypted save"));
      ExcelOoxmlPackageSecuritySupport.saveWorkbook(
          workbook,
          encryptedOutput,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          new ExcelOoxmlPersistenceOptions(
              Optional.of(
                  new ExcelOoxmlEncryptionOptions(
                      "secret-password",
                      ExcelOoxmlWriteCipher.AES_256,
                      ExcelOoxmlWriteHash.SHA_512)),
              Optional.empty()),
          (prefix, suffix) -> {
            tempFilesCreated.incrementAndGet();
            return ExcelTempFiles.createManagedTempFile(explicitTempRoot, prefix, suffix);
          });
    }

    assertEquals(0, tempFilesCreated.get());
    assertTrue(Files.notExists(explicitTempRoot));
    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(
            encryptedOutput,
            new ExcelOoxmlOpenOptions.Encrypted("secret-password"),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Encrypted save", reopened.sheet("Encrypted").cells().text("A1"));
    }
  }

  private static void assertCopyDeleteAndEffectiveOptionsBranches(Path sourceWorkbookPath)
      throws IOException {
    ExcelOoxmlPackageFileSupport.copySourceWorkbook(
        sourceWorkbookPath, sourceWorkbookPath, WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
    assertTrue(OoxmlSecurityTestSupport.signatureValid(sourceWorkbookPath));

    Path copiedWorkbookPath = sourceWorkbookPath.getParent().resolve("copied-signed.xlsx");
    ExcelOoxmlPackageFileSupport.copySourceWorkbook(
        sourceWorkbookPath, copiedWorkbookPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
    assertArrayEquals(
        Files.readAllBytes(sourceWorkbookPath), Files.readAllBytes(copiedWorkbookPath));

    Path deletedTempFile =
        ExcelTempFiles.createManagedTempFile("gridgrind-delete-if-exists-", ".tmp");
    ExcelOoxmlPackageFileSupport.deleteIfExists(null);
    ExcelOoxmlPackageFileSupport.deleteIfExists(deletedTempFile);
    assertFalse(Files.exists(deletedTempFile));

    Path nonEmptyDirectory =
        ExcelTempFiles.createManagedTempDirectory("gridgrind-delete-if-exists-dir-");
    Files.writeString(nonEmptyDirectory.resolve("child.txt"), "keep");
    ExcelOoxmlPackageFileSupport.deleteIfExists(nonEmptyDirectory);
    assertTrue(Files.exists(nonEmptyDirectory));
    ExcelOoxmlPackageFileSupport.deleteTreeIfExists(null);
    ExcelOoxmlPackageFileSupport.deleteTreeIfExists(
        nonEmptyDirectory.resolve("missing-cleanup-root"));
    ExcelOoxmlPackageFileSupport.deleteTreeIfExists(nonEmptyDirectory);
    assertFalse(Files.exists(nonEmptyDirectory));

    ExcelOoxmlEncryptionSnapshot encryptedSnapshot =
        new ExcelOoxmlEncryptionSnapshot.Encrypted(
            ExcelOoxmlEncryptionMode.AGILE,
            ExcelOoxmlCipherAlgorithm.AES_256,
            ExcelOoxmlHashAlgorithm.SHA_512,
            ExcelOoxmlChainingMode.CBC,
            256,
            16,
            100_000);
    IllegalStateException missingPasswordFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelOoxmlPackagePersistenceSupport.effectiveOptions(
                    new ExcelOoxmlPackageSecuritySnapshot(encryptedSnapshot, java.util.List.of()),
                    Optional.empty(),
                    ExcelOoxmlPersistenceOptions.none()));
    assertTrue(missingPasswordFailure.getMessage().contains("verified source password"));

    ExcelOoxmlPersistenceOptions preservedOptions =
        ExcelOoxmlPackagePersistenceSupport.effectiveOptions(
            new ExcelOoxmlPackageSecuritySnapshot(encryptedSnapshot, java.util.List.of()),
            Optional.of("persist-pass"),
            ExcelOoxmlPersistenceOptions.none());
    assertEquals(
        ExcelOoxmlWriteCipher.AES_256, preservedOptions.encryption().orElseThrow().cipher());
    assertEquals(ExcelOoxmlWriteHash.SHA_512, preservedOptions.encryption().orElseThrow().hash());

    ExcelOoxmlEncryptionSnapshot standardSnapshot =
        new ExcelOoxmlEncryptionSnapshot.Encrypted(
            ExcelOoxmlEncryptionMode.STANDARD,
            ExcelOoxmlCipherAlgorithm.AES_128,
            ExcelOoxmlHashAlgorithm.SHA_1,
            ExcelOoxmlChainingMode.ECB,
            128,
            16,
            50_000);
    IllegalArgumentException standardPreservationFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelOoxmlPackagePersistenceSupport.effectiveOptions(
                    new ExcelOoxmlPackageSecuritySnapshot(standardSnapshot, java.util.List.of()),
                    Optional.of("persist-pass"),
                    ExcelOoxmlPersistenceOptions.none()));
    assertTrue(standardPreservationFailure.getMessage().contains("not auto-preservable"));

    ExcelOoxmlEncryptionSnapshot nonCbcSnapshot =
        new ExcelOoxmlEncryptionSnapshot.Encrypted(
            ExcelOoxmlEncryptionMode.AGILE,
            ExcelOoxmlCipherAlgorithm.AES_256,
            ExcelOoxmlHashAlgorithm.SHA_512,
            ExcelOoxmlChainingMode.CFB,
            256,
            16,
            100_000);
    IllegalArgumentException nonCbcPreservationFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelOoxmlPackagePersistenceSupport.effectiveOptions(
                    new ExcelOoxmlPackageSecuritySnapshot(nonCbcSnapshot, java.util.List.of()),
                    Optional.of("persist-pass"),
                    ExcelOoxmlPersistenceOptions.none()));
    assertTrue(nonCbcPreservationFailure.getMessage().contains("chaining mode"));
    assertTrue(nonCbcPreservationFailure.getMessage().contains("CFB"));

    ExcelOoxmlEncryptionSnapshot unsupportedCipherSnapshot =
        new ExcelOoxmlEncryptionSnapshot.Encrypted(
            ExcelOoxmlEncryptionMode.AGILE,
            ExcelOoxmlCipherAlgorithm.AES_128,
            ExcelOoxmlHashAlgorithm.SHA_512,
            ExcelOoxmlChainingMode.CBC,
            128,
            16,
            100_000);
    IllegalArgumentException unsupportedCipherPreservationFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelOoxmlPackagePersistenceSupport.effectiveOptions(
                    new ExcelOoxmlPackageSecuritySnapshot(
                        unsupportedCipherSnapshot, java.util.List.of()),
                    Optional.of("persist-pass"),
                    ExcelOoxmlPersistenceOptions.none()));
    assertTrue(unsupportedCipherPreservationFailure.getMessage().contains("cipher AES_128"));

    ExcelOoxmlEncryptionSnapshot unsupportedHashSnapshot =
        new ExcelOoxmlEncryptionSnapshot.Encrypted(
            ExcelOoxmlEncryptionMode.AGILE,
            ExcelOoxmlCipherAlgorithm.AES_256,
            ExcelOoxmlHashAlgorithm.SHA_1,
            ExcelOoxmlChainingMode.CBC,
            256,
            16,
            100_000);
    IllegalArgumentException unsupportedHashPreservationFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                ExcelOoxmlPackagePersistenceSupport.effectiveOptions(
                    new ExcelOoxmlPackageSecuritySnapshot(
                        unsupportedHashSnapshot, java.util.List.of()),
                    Optional.of("persist-pass"),
                    ExcelOoxmlPersistenceOptions.none()));
    assertTrue(unsupportedHashPreservationFailure.getMessage().contains("hash SHA_1"));
  }

  private static void assertSigningMaterialAndAliasBranches(
      OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook,
      KeyStore signingKeyStore,
      ExcelOoxmlSignatureOptions signatureOptions)
      throws IOException, GeneralSecurityException {
    assertNotNull(ExcelOoxmlPackageSigningSupport.signingMaterial(signatureOptions));

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
                ExcelOoxmlPackageSigningSupport.loadSigningKeyStore(
                    signedWorkbook.pkcs12Path(), wrongKeystorePassword));
    assertTrue(
        wrongKeystorePasswordFailure.getMessage().contains("Failed to load signing material"));

    InvalidSigningConfigurationException missingSigningMaterialFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.signingMaterial(
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
        ExcelOoxmlPackageSigningSupport.resolveAlias(
            signingKeyStore, signedWorkbook.alias(), signatureOptions));

    InvalidSigningConfigurationException missingAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.resolveAlias(
                    signingKeyStore, "missing-alias", signatureOptions));
    assertTrue(missingAliasFailure.getMessage().contains("does not exist"));

    KeyStore emptyKeyStore = KeyStore.getInstance("PKCS12");
    emptyKeyStore.load(null, null);
    InvalidSigningConfigurationException noKeyAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.resolveAlias(
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
                ExcelOoxmlPackageSigningSupport.signingPrivateKey(
                    certificateOnlyKeyStore,
                    signedWorkbook.pkcs12Path(),
                    "certificate-only",
                    signatureOptions));
    assertTrue(nonPrivateKeyFailure.getMessage().contains("does not resolve to a private key"));

    InvalidSigningConfigurationException certificateChainFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.signingCertificateChain(
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
                ExcelOoxmlPackageSigningSupport.resolveSigningAlias(
                    uninitializedKeyStore, signedWorkbook.pkcs12Path(), signatureOptions));
    assertTrue(aliasInspectionFailure.getMessage().contains("Failed to inspect signing aliases"));

    InvalidSigningConfigurationException privateKeyLoadFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.signingPrivateKey(
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
                ExcelOoxmlPackageSigningSupport.signingCertificateChain(
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
                ExcelOoxmlPackageSigningSupport.resolveAlias(
                    multiAliasKeyStore, null, signatureOptions));
    assertTrue(multipleAliasesFailure.getMessage().contains("multiple private-key aliases"));

    KeyStore certificateOnlyAliasKeyStore = KeyStore.getInstance("PKCS12");
    certificateOnlyAliasKeyStore.load(null, null);
    certificateOnlyAliasKeyStore.setCertificateEntry("certificate-only", certificate);
    InvalidSigningConfigurationException certificateOnlyAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.resolveAlias(
                    certificateOnlyAliasKeyStore, null, signatureOptions));
    assertTrue(
        certificateOnlyAliasFailure.getMessage().contains("does not contain a private-key entry"));

    KeyStore nonPrivateKeyStore = fakeKeyStore(certificate.getPublicKey(), null);
    InvalidSigningConfigurationException nonPrivateAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.resolveAlias(
                    nonPrivateKeyStore, null, signatureOptions));
    assertTrue(
        nonPrivateAliasFailure.getMessage().contains("does not contain a private-key entry"));
    InvalidSigningConfigurationException reflectedNonPrivateKeyFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                ExcelOoxmlPackageSigningSupport.signingPrivateKey(
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
                ExcelOoxmlPackageSigningSupport.signingCertificateChain(
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
                ExcelOoxmlPackageSigningSupport.signingCertificateChain(
                    emptyChainKeyStore, signedWorkbook.pkcs12Path(), "alias"));
    assertTrue(emptyChainFailure.getMessage().contains("X.509 certificate chain"));

    Path signableWorkbookPath =
        ExcelTempFiles.createManagedTempFile("gridgrind-sign-description-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Signed")
          .cells()
          .setCell("A1", ExcelCellValue.text("Signed workbook"));
      workbook
          .persistence()
          .save(
              signableWorkbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }
    ExcelOoxmlPackageSigningSupport.signWorkbook(
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
