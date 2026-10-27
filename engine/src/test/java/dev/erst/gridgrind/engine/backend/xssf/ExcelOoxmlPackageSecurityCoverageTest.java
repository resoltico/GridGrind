package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment;
import dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureReport;
import dev.erst.gridgrind.core.model.security.ResolvedEncryption;
import dev.erst.gridgrind.core.model.security.ResolvedOpenSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceEncryption;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSignature;
import dev.erst.gridgrind.core.model.security.ResolvedSignature;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.engine.packaging.WorkbookFiles;
import dev.erst.gridgrind.engine.security.SecurityAdapter;
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
    Path directory = ManagedTempFiles.createManagedTempDirectory("gridgrind-ooxml-materialize-");
    Path plainWorkbookPath = directory.resolve("plain.xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Plain")
          .cells()
          .setCell("A1", new CellValue.Text("Plain workbook"));
      workbook
          .persistence()
          .save(
              plainWorkbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (SecurityAdapter.ReadableWorkbook readableWorkbook =
        SecurityAdapter.materializeReadableWorkbook(
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
          return ManagedTempFiles.createManagedTempFile(explicitTempRoot, prefix, suffix);
        };

    try (SecurityAdapter.ReadableWorkbook readableWorkbook =
        SecurityAdapter.materializeReadableWorkbook(
            encryptedWorkbook.workbookPath(),
            new ResolvedOpenSecurity.Encrypted(encryptedWorkbook.password()),
            tempFileFactory)) {
      assertNotEquals(
          encryptedWorkbook.workbookPath().toAbsolutePath().normalize(),
          readableWorkbook.workbookPath());
      assertEquals(
          Optional.of(encryptedWorkbook.password()), readableWorkbook.sourceEncryptionPassword());
      assertInstanceOf(
          OoxmlEncryptionReport.Encrypted.class, readableWorkbook.packageSecurity().encryption());
      assertTrue(Files.exists(readableWorkbook.workbookPath()));
      assertFalse(
          readableWorkbook
              .workbookPath()
              .startsWith(explicitTempRoot.toAbsolutePath().normalize()));
    }
    assertEquals(0, tempFilesCreated.get());
    assertTrue(Files.notExists(explicitTempRoot));

    Path legacyWorkbookPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-legacy-materialize-", ".xls");
    try (HSSFWorkbook workbook = new HSSFWorkbook();
        OutputStream outputStream = Files.newOutputStream(legacyWorkbookPath)) {
      workbook.createSheet("Legacy").createRow(0).createCell(0).setCellValue("Legacy workbook");
      workbook.write(outputStream);
    }

    WorkbookNotOpenableException legacyFailure =
        assertThrows(
            WorkbookNotOpenableException.class,
            () ->
                SecurityAdapter.materializeReadableWorkbook(
                    legacyWorkbookPath,
                    new ResolvedOpenSecurity.Encrypted("unused"),
                    Files::createTempFile));
    assertTrue(legacyFailure.getMessage().contains("not openable"));
  }

  @Test
  void workbookPackageIoBridgesMaterializationAndPersistenceForPlainWorkbooks() throws IOException {
    Path directory = ManagedTempFiles.createManagedTempDirectory("gridgrind-package-io-");
    Path sourceWorkbookPath = directory.resolve("source.xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Bridge"));
      workbook
          .persistence()
          .save(
              sourceWorkbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    Path copiedWorkbookPath = directory.resolve("copied.xlsx");
    try (WorkbookArtifactIo.MaterializedWorkbook materializedWorkbook =
        WorkbookArtifactIo.materializeWorkbook(sourceWorkbookPath, null, Files::createTempFile)) {
      assertEquals(
          sourceWorkbookPath.toAbsolutePath().normalize(), materializedWorkbook.workbookPath());
      WorkbookArtifactIo.persistMaterializedWorkbook(
          materializedWorkbook.workbookPath(),
          copiedWorkbookPath,
          OoxmlPackageSecurityReport.none(),
          Optional.empty(),
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          ResolvedPersistenceSecurity.none());
    }

    try (XssfWorkbook reopened =
        XssfWorkbooks.open(copiedWorkbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Bridge", reopened.sheet("Budget").cells().text("A1"));
    }
  }

  @Test
  void workbookFileSupportCoversReplaceExistingCopyAndMoveBranches() throws IOException {
    Path workspace = ManagedTempFiles.createManagedTempDirectory("gridgrind-package-file-ops-");
    Path sourceWorkbookPath = workspace.resolve("source.xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", new CellValue.Text("Source"));
      workbook
          .persistence()
          .save(
              sourceWorkbookPath,
              WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    Path replaceExistingTarget = workspace.resolve("replace-existing.xlsx");
    Files.writeString(replaceExistingTarget, "old");
    WorkbookFiles.copySourceWorkbook(
        sourceWorkbookPath,
        replaceExistingTarget,
        WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
    assertArrayEquals(
        Files.readAllBytes(sourceWorkbookPath), Files.readAllBytes(replaceExistingTarget));
    NullPointerException nullTargetFailure =
        assertThrows(
            NullPointerException.class,
            () ->
                WorkbookFiles.copySourceWorkbook(
                    sourceWorkbookPath, null, WorkbookArtifactWriteDisposition.CREATE_NEW));
    assertEquals("targetPath must not be null", nullTargetFailure.getMessage());
    WorkbookFiles.copySourceWorkbook(
        sourceWorkbookPath, sourceWorkbookPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
    assertTrue(Files.exists(sourceWorkbookPath));

    Path moveCreateNewSource = workspace.resolve("move-create-new-source.xlsx");
    Files.copy(sourceWorkbookPath, moveCreateNewSource);
    Path moveCreateNewTarget = workspace.resolve("move-create-new-target.xlsx");
    WorkbookFiles.moveWorkbook(
        moveCreateNewSource, moveCreateNewTarget, WorkbookArtifactWriteDisposition.CREATE_NEW);
    assertFalse(Files.exists(moveCreateNewSource));
    assertTrue(Files.exists(moveCreateNewTarget));

    Path moveReplaceSource = workspace.resolve("move-replace-source.xlsx");
    Files.copy(sourceWorkbookPath, moveReplaceSource);
    Path moveReplaceTarget = workspace.resolve("move-replace-target.xlsx");
    Files.writeString(moveReplaceTarget, "old target");
    WorkbookFiles.moveWorkbook(
        moveReplaceSource, moveReplaceTarget, WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
    assertFalse(Files.exists(moveReplaceSource));
    assertArrayEquals(
        Files.readAllBytes(sourceWorkbookPath), Files.readAllBytes(moveReplaceTarget));

    Path noOpMovePath = workspace.resolve("move-no-op.xlsx");
    Files.copy(sourceWorkbookPath, noOpMovePath);
    WorkbookFiles.moveWorkbook(
        noOpMovePath, noOpMovePath, WorkbookArtifactWriteDisposition.CREATE_NEW);
    assertTrue(Files.exists(noOpMovePath));
  }

  @Test
  void encryptedMaterializationCleansTempFilesAfterPasswordFailure() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            ManagedTempFiles.createManagedTempDirectory("gridgrind-ooxml-materialize-cleanup-"));
    Path explicitTempRoot = encryptedWorkbook.workbookPath().getParent().resolve("caller-scratch");
    AtomicInteger tempFilesCreated = new AtomicInteger();

    InvalidWorkbookPasswordException failure =
        assertThrows(
            InvalidWorkbookPasswordException.class,
            () ->
                SecurityAdapter.materializeReadableWorkbook(
                    encryptedWorkbook.workbookPath(),
                    new ResolvedOpenSecurity.Encrypted("wrong-password"),
                    (prefix, suffix) -> {
                      tempFilesCreated.incrementAndGet();
                      return ManagedTempFiles.createManagedTempFile(
                          explicitTempRoot, prefix, suffix);
                    }));

    assertFalse(failure.getMessage().contains(encryptedWorkbook.workbookPath().toString()));
    assertEquals(0, tempFilesCreated.get());
    assertTrue(Files.notExists(explicitTempRoot));
  }

  @Test
  void encryptedMaterializationRequiresOpenOptionsPassword() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            ManagedTempFiles.createManagedTempDirectory("gridgrind-ooxml-password-required-"));

    WorkbookPasswordRequiredException failure =
        assertThrows(
            WorkbookPasswordRequiredException.class,
            () ->
                SecurityAdapter.materializeReadableWorkbook(
                    encryptedWorkbook.workbookPath(), null, Files::createTempFile));

    assertFalse(failure.getMessage().contains(encryptedWorkbook.workbookPath().toString()));

    WorkbookPasswordRequiredException nullPasswordFailure =
        assertThrows(
            WorkbookPasswordRequiredException.class,
            () ->
                SecurityAdapter.materializeReadableWorkbook(
                    encryptedWorkbook.workbookPath(),
                    new ResolvedOpenSecurity.Unencrypted(),
                    Files::createTempFile));
    assertFalse(
        nullPasswordFailure.getMessage().contains(encryptedWorkbook.workbookPath().toString()));
  }

  @Test
  void unsupportedMagicAndMaterializedWorkbookEdgeCasesStayExplicit() throws IOException {
    Path unsupportedPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-unsupported-magic-", ".bin");
    Files.writeString(unsupportedPath, "plain text is not an OOXML workbook");
    WorkbookNotOpenableException unsupportedMagicFailure =
        assertThrows(
            WorkbookNotOpenableException.class,
            () ->
                SecurityAdapter.materializeReadableWorkbook(
                    unsupportedPath, null, Files::createTempFile));
    assertTrue(unsupportedMagicFailure.getMessage().contains("not openable"));

    Path missingWorkbookPath =
        ManagedTempFiles.createManagedTempDirectory("gridgrind-materialized-missing-")
            .resolve("missing.xlsx");
    WorkbookNotFoundException missingPlainFailure =
        assertThrows(
            WorkbookNotFoundException.class,
            () ->
                XssfWorkbooks.openMaterializedWorkbook(
                    missingWorkbookPath,
                    Optional.of(missingWorkbookPath),
                    OoxmlPackageSecurityReport.none(),
                    Optional.empty()));
    assertEquals(missingWorkbookPath.toAbsolutePath(), missingPlainFailure.workbookPath());

    WorkbookNotFoundException missingFormulaFailure =
        assertThrows(
            WorkbookNotFoundException.class,
            () ->
                XssfWorkbooks.openMaterializedWorkbook(
                    missingWorkbookPath,
                    ResolvedFormulaEnvironment.defaults(),
                    Optional.of(missingWorkbookPath),
                    OoxmlPackageSecurityReport.none(),
                    Optional.empty()));
    assertEquals(missingWorkbookPath.toAbsolutePath(), missingFormulaFailure.workbookPath());

    Path nonXlsxPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-materialized-legacy-", ".xls");
    try (HSSFWorkbook workbook = new HSSFWorkbook();
        OutputStream outputStream = Files.newOutputStream(nonXlsxPath)) {
      workbook.createSheet("Legacy").createRow(0).createCell(0).setCellValue("Legacy workbook");
      workbook.write(outputStream);
    }

    IllegalArgumentException unsupportedMaterializedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                XssfWorkbooks.openMaterializedWorkbook(
                    nonXlsxPath,
                    Optional.of(nonXlsxPath),
                    OoxmlPackageSecurityReport.none(),
                    Optional.empty()));
    assertEquals("Only .xlsx workbooks are supported", unsupportedMaterializedFailure.getMessage());

    IllegalArgumentException unsupportedFormulaMaterializedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                XssfWorkbooks.openMaterializedWorkbook(
                    nonXlsxPath,
                    ResolvedFormulaEnvironment.defaults(),
                    Optional.of(nonXlsxPath),
                    OoxmlPackageSecurityReport.none(),
                    Optional.empty()));
    assertEquals(
        "Only .xlsx workbooks are supported", unsupportedFormulaMaterializedFailure.getMessage());
  }

  @Test
  void workbookOpenOverloadsCoverCustomTempFactoriesAndMaterializedPaths() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            ManagedTempFiles.createManagedTempDirectory("gridgrind-workbook-open-security-"));
    AtomicInteger tempFilesCreated = new AtomicInteger();
    WorkbookTempFileFactory tempFileFactory =
        (prefix, suffix) -> {
          tempFilesCreated.incrementAndGet();
          return ManagedTempFiles.createManagedTempFile(prefix, suffix);
        };

    try (XssfWorkbook workbook =
            XssfWorkbooks.open(
                encryptedWorkbook.workbookPath(),
                new ResolvedOpenSecurity.Encrypted(encryptedWorkbook.password()),
                tempFileFactory);
        XssfWorkbook workbookWithEnvironment =
            XssfWorkbooks.open(
                encryptedWorkbook.workbookPath(),
                ResolvedFormulaEnvironment.defaults(),
                new ResolvedOpenSecurity.Encrypted(encryptedWorkbook.password()),
                tempFileFactory)) {
      assertEquals("Encrypted workbook", workbook.sheet("Encrypted").cells().text("A1"));
      assertEquals(
          "Encrypted workbook", workbookWithEnvironment.sheet("Encrypted").cells().text("A1"));
    }
    assertEquals(0, tempFilesCreated.get());

    Path materializedWorkbookPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-materialized-open-", ".xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Plain").cells().setCell("A1", new CellValue.Text("Materialized"));
      workbook
          .persistence()
          .save(
              materializedWorkbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XssfWorkbook workbook =
            XssfWorkbooks.openMaterializedWorkbook(
                materializedWorkbookPath,
                Optional.of(materializedWorkbookPath),
                OoxmlPackageSecurityReport.none(),
                Optional.empty());
        XssfWorkbook workbookWithEnvironment =
            XssfWorkbooks.openMaterializedWorkbook(
                materializedWorkbookPath,
                ResolvedFormulaEnvironment.defaults(),
                Optional.of(materializedWorkbookPath),
                OoxmlPackageSecurityReport.none(),
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
                XssfWorkbooks.openMaterializedWorkbook(
                    new XSSFWorkbook(), Optional.empty(), null, Optional.empty()));
    assertEquals("loadedPackageSecurity must not be null", noFormulaFailure.getMessage());

    try (ThrowingOpenCloseWorkbook throwingNoFormulaWorkbook =
        new ThrowingOpenCloseWorkbook("no-formula close failure")) {
      try {
        NullPointerException throwingNoFormulaFailure =
            assertThrows(
                NullPointerException.class,
                () ->
                    XssfWorkbooks.openMaterializedWorkbook(
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
                    XssfWorkbooks.openMaterializedWorkbook(
                        throwingFormulaWorkbook,
                        ResolvedFormulaEnvironment.defaults(),
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
      XssfWorkbooks.closeWorkbookAfterOpenFailure(workbook, runtimeFailure);
      assertEquals(0, runtimeFailure.getSuppressed().length);
    }
  }

  @Test
  void securityHelpersCoverRemainingAliasAndFilesystemBranches() throws Exception {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            ManagedTempFiles.createManagedTempDirectory("gridgrind-signing-helper-branches-"));
    KeyStore signingKeyStore =
        loadPkcs12(signedWorkbook.pkcs12Path(), signedWorkbook.keystorePassword());
    Certificate certificate = signingKeyStore.getCertificate(signedWorkbook.alias());

    assertCopyDeleteAndEffectiveOptionsBranches(signedWorkbook.workbookPath());

    ResolvedSignature signatureOptions =
        new ResolvedSignature(
            signedWorkbook.pkcs12Path(),
            signedWorkbook.keystorePassword(),
            signedWorkbook.keyPassword(),
            java.util.Optional.of(signedWorkbook.alias()),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            java.util.Optional.empty());
    assertSigningMaterialAndAliasBranches(signedWorkbook, signingKeyStore, signatureOptions);
    assertCertificateOnlyAndUninitializedKeystoreBranches(
        signedWorkbook, certificate, signatureOptions);
    assertSyntheticKeystoreBranches(signedWorkbook, signingKeyStore, certificate, signatureOptions);
  }

  @Test
  void failureTranslationHelpersCoverSecurityBridgeDefensivePaths() throws Exception {
    Path invalidEncryptedWorkbookPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-invalid-encrypted-ooxml-", ".xlsx");
    try (POIFSFileSystem fileSystem = new POIFSFileSystem()) {
      fileSystem.createDocument(
          new ByteArrayInputStream(new byte[0]), Decryptor.DEFAULT_POIFS_ENTRY);
      try (OutputStream outputStream = Files.newOutputStream(invalidEncryptedWorkbookPath)) {
        fileSystem.writeFilesystem(outputStream);
      }
    }
    try (POIFSFileSystem invalidFileSystem =
        new POIFSFileSystem(invalidEncryptedWorkbookPath.toFile())) {
      WorkbookNotOpenableException invalidEncryptionInfoFailure =
          assertThrows(
              WorkbookNotOpenableException.class,
              () ->
                  SecurityAdapter.readEncryptionInfo(
                      invalidFileSystem, invalidEncryptedWorkbookPath));
      assertTrue(invalidEncryptionInfoFailure.getMessage().contains("not openable"));
    }

    WorkbookSecurityException passwordFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                SecurityAdapter.verifyPassword(
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
                SecurityAdapter.materializeDecryptedWorkbook(
                    () -> {
                      throw new GeneralSecurityException("decrypt failure");
                    },
                    ManagedTempFiles.createManagedTempFile("gridgrind-decrypt-failure-", ".xlsx"),
                    invalidEncryptedWorkbookPath));
    assertTrue(
        decryptFailure.getMessage().contains("Failed to decrypt the OOXML workbook package"));

    EncryptionInfo brokenEncryptionInfo = new EncryptionInfo(EncryptionMode.agile);
    brokenEncryptionInfo.setHeader(null);
    WorkbookSecurityException encryptionSnapshotFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () -> SecurityAdapter.encryptionSnapshot(brokenEncryptionInfo));
    assertTrue(
        encryptionSnapshotFailure.getMessage().contains("inspect OOXML encryption metadata"));

    Path invalidPackagePath =
        ManagedTempFiles.createManagedTempFile("gridgrind-invalid-package-", ".txt");
    Files.writeString(invalidPackagePath, "not a package");
    WorkbookNotOpenableException packageInspectionFailure =
        assertThrows(
            WorkbookNotOpenableException.class,
            () ->
                SecurityAdapter.inspectPackageSecurity(
                    invalidPackagePath, new OoxmlEncryptionReport.None()));
    assertTrue(packageInspectionFailure.getMessage().contains("not openable"));

    assertEquals(Optional.empty(), SecurityAdapter.signerIdentity(null));

    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            ManagedTempFiles.createManagedTempDirectory("gridgrind-signer-metadata-"));
    KeyStore signingKeyStore =
        loadPkcs12(signedWorkbook.pkcs12Path(), signedWorkbook.keystorePassword());
    java.security.cert.X509Certificate signer =
        (java.security.cert.X509Certificate) signingKeyStore.getCertificate(signedWorkbook.alias());
    OoxmlSignatureReport.SignerIdentity signerIdentity =
        SecurityAdapter.signerIdentity(signer).orElseThrow();
    assertTrue(signerIdentity.subject().contains("GridGrind Signing Test"));
    assertTrue(signerIdentity.issuer().contains("GridGrind Signing Test"));
    assertFalse(signerIdentity.serialNumberHex().isBlank());

    OoxmlSignatureReport inspectedSignature =
        SecurityAdapter.inspectPackageSecurity(
                signedWorkbook.workbookPath(), new OoxmlEncryptionReport.None())
            .signatures()
            .getFirst();
    assertEquals(ExcelOoxmlSignatureState.VALID, inspectedSignature.state());
    assertTrue(
        inspectedSignature.signer().orElseThrow().subject().contains("GridGrind Signing Test"));

    WorkbookSecurityException signingFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                SecurityAdapter.confirmAndVerifySignature(
                    () -> {
                      throw new javax.xml.crypto.dsig.XMLSignatureException("sign failure");
                    },
                    invalidEncryptedWorkbookPath));
    assertTrue(signingFailure.getMessage().contains("Failed to sign the OOXML workbook package"));

    WorkbookSecurityException runtimeSigningFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                SecurityAdapter.confirmAndVerifySignature(
                    () -> {
                      throw new IllegalStateException("runtime sign failure");
                    },
                    invalidEncryptedWorkbookPath));
    assertTrue(runtimeSigningFailure.getMessage().contains("Unexpected OOXML signing failure"));

    WorkbookSecurityException invalidSignatureFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                SecurityAdapter.confirmAndVerifySignature(
                    () -> false, invalidEncryptedWorkbookPath));
    assertTrue(invalidSignatureFailure.getMessage().contains("did not validate after signing"));

    WorkbookSecurityException encryptFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () ->
                SecurityAdapter.writeEncryptedWorkbook(
                    fileSystem -> {
                      throw new GeneralSecurityException("encrypt failure");
                    },
                    ManagedTempFiles.createManagedTempFile("gridgrind-encrypt-source-", ".xlsx"),
                    ManagedTempFiles.createManagedTempFile("gridgrind-encrypt-target-", ".xlsx"),
                    WorkbookArtifactWriteDisposition.CREATE_NEW));
    assertTrue(
        encryptFailure.getMessage().contains("Failed to encrypt the saved OOXML workbook package"));

    ResolvedSignature signatureOptions =
        new ResolvedSignature(
            signedWorkbook.pkcs12Path(),
            signedWorkbook.keystorePassword(),
            signedWorkbook.keyPassword(),
            java.util.Optional.empty(),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            java.util.Optional.empty());
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
        "only-alias", SecurityAdapter.resolveAlias(singleAliasKeyStore, null, signatureOptions));

    Path malformedPackagePath =
        ManagedTempFiles.createManagedTempFile("gridgrind-sign-invalid-format-", ".xlsx");
    try (var outputStream =
        new java.util.zip.ZipOutputStream(Files.newOutputStream(malformedPackagePath))) {
      outputStream.putNextEntry(new java.util.zip.ZipEntry("broken.txt"));
      outputStream.write("broken".getBytes(java.nio.charset.StandardCharsets.UTF_8));
      outputStream.closeEntry();
    }
    WorkbookSecurityException invalidFormatFailure =
        assertThrows(
            WorkbookSecurityException.class,
            () -> SecurityAdapter.signWorkbook(malformedPackagePath, signatureOptions));
    assertTrue(invalidFormatFailure.getMessage().contains("open the OOXML workbook package"));
  }

  @Test
  void saveAndPersistSecurityHelpersCoverPlainAndSignedGuardBranches() throws IOException {
    Path workspace =
        ManagedTempFiles.createManagedTempDirectory("gridgrind-saveworkbook-security-");
    Path plainSourcePath = workspace.resolve("plain-source.xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Plain").cells().setCell("A1", new CellValue.Text("Plain save"));
      workbook
          .persistence()
          .save(
              plainSourcePath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    Path plainSavedPath = workspace.resolve("plain-saved-via-support.xlsx");
    try (XssfWorkbook workbook =
        XssfWorkbooks.open(plainSourcePath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      SecurityAdapter.saveWorkbook(
          workbook,
          plainSavedPath,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          ResolvedPersistenceSecurity.none(),
          Files::createTempFile);
    }
    try (XssfWorkbook workbook =
        XssfWorkbooks.open(plainSavedPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Plain save", workbook.sheet("Plain").cells().text("A1"));
    }

    Path plainWorkbookPath = workspace.resolve("persist-signed-plain.xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Guard").cells().setCell("A1", new CellValue.Text("Signed guard"));
      workbook
          .persistence()
          .save(
              plainWorkbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    OoxmlSignatureReport signature =
        new OoxmlSignatureReport(
            "/_xmlsignatures/sig1.xml",
            Optional.of("CN=GridGrind Signing Test"),
            Optional.of("CN=GridGrind Signing Test"),
            Optional.of("01AB"),
            ExcelOoxmlSignatureState.VALID);
    OoxmlPackageSecurityReport signedSecurity =
        new OoxmlPackageSecurityReport(
            new OoxmlEncryptionReport.None(), java.util.List.of(signature));

    Path unsignedTarget = plainWorkbookPath.resolveSibling("signed-guard-output.xlsx");
    SecurityAdapter.persistMaterializedWorkbook(
        plainWorkbookPath,
        unsignedTarget,
        signedSecurity,
        Optional.empty(),
        WorkbookArtifactWriteDisposition.CREATE_NEW,
        ResolvedPersistenceSecurity.none());
    assertTrue(Files.exists(unsignedTarget));
  }

  @Test
  void saveWorkbookCoversInMemoryAndSignedPassthroughBranches() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Memory").cells().setCell("A1", new CellValue.Text("Memory save"));
      Path memoryTarget = ManagedTempFiles.createManagedTempFile("gridgrind-save-memory-", ".xlsx");

      SecurityAdapter.saveWorkbook(
          workbook,
          memoryTarget,
          WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
          null,
          Files::createTempFile);

      try (XssfWorkbook reopened =
          XssfWorkbooks.open(memoryTarget, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
        assertEquals("Memory save", reopened.sheet("Memory").cells().text("A1"));
      }
    }

    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            ManagedTempFiles.createManagedTempDirectory("gridgrind-save-signed-passthrough-"));
    Path copiedWorkbookPath =
        signedWorkbook.workbookPath().getParent().resolve("copied-signed.xlsx");
    AtomicInteger tempFilesCreated = new AtomicInteger();
    try (XssfWorkbook workbook =
        XssfWorkbooks.open(
            signedWorkbook.workbookPath(), ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      SecurityAdapter.saveWorkbook(
          workbook,
          copiedWorkbookPath,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          ResolvedPersistenceSecurity.none(),
          (prefix, suffix) -> {
            tempFilesCreated.incrementAndGet();
            return ManagedTempFiles.createManagedTempFile(prefix, suffix);
          });
    }
    assertEquals(0, tempFilesCreated.get());
    assertTrue(
        SecurityAdapter.inspectPackageSecurity(copiedWorkbookPath, new OoxmlEncryptionReport.None())
            .signatures()
            .isEmpty());

    Path resignedWorkbookPath = signedWorkbook.workbookPath().getParent().resolve("resigned.xlsx");
    try (XssfWorkbook workbook =
        XssfWorkbooks.open(
            signedWorkbook.workbookPath(), ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      SecurityAdapter.saveWorkbook(
          workbook,
          resignedWorkbookPath,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          new ResolvedPersistenceSecurity(
              new OoxmlPersistenceEncryptionSpec.None(),
              new ResolvedPersistenceSignature.Sign(
                  new ResolvedSignature(
                      signedWorkbook.pkcs12Path(),
                      signedWorkbook.keystorePassword(),
                      signedWorkbook.keyPassword(),
                      java.util.Optional.of(signedWorkbook.alias()),
                      ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                      java.util.Optional.empty()))),
          Files::createTempFile);
    }
    assertTrue(OoxmlSecurityTestSupport.signatureValid(resignedWorkbookPath));
  }

  @Test
  void encryptedSaveKeepsPlaintextTempsOutOfCallerScratchRoot() throws IOException {
    Path workspace =
        ManagedTempFiles.createManagedTempDirectory("gridgrind-encrypted-save-private-");
    Path explicitTempRoot = workspace.resolve("caller-scratch");
    Path encryptedOutput = workspace.resolve("encrypted-output.xlsx");
    AtomicInteger tempFilesCreated = new AtomicInteger();

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Encrypted")
          .cells()
          .setCell("A1", new CellValue.Text("Encrypted save"));
      SecurityAdapter.saveWorkbook(
          workbook,
          encryptedOutput,
          WorkbookArtifactWriteDisposition.CREATE_NEW,
          new ResolvedPersistenceSecurity(
              new ResolvedPersistenceEncryption.Encrypt(
                  new ResolvedEncryption(
                      "secret-password",
                      ExcelOoxmlWriteCipher.AES_256,
                      ExcelOoxmlWriteHash.SHA_512)),
              new OoxmlPersistenceSignatureSpec.None()),
          (prefix, suffix) -> {
            tempFilesCreated.incrementAndGet();
            return ManagedTempFiles.createManagedTempFile(explicitTempRoot, prefix, suffix);
          });
    }

    assertEquals(0, tempFilesCreated.get());
    assertTrue(Files.notExists(explicitTempRoot));
    try (XssfWorkbook reopened =
        XssfWorkbooks.open(
            encryptedOutput,
            new ResolvedOpenSecurity.Encrypted("secret-password"),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Encrypted save", reopened.sheet("Encrypted").cells().text("A1"));
    }
  }

  private static void assertCopyDeleteAndEffectiveOptionsBranches(Path sourceWorkbookPath)
      throws IOException {
    WorkbookFiles.copySourceWorkbook(
        sourceWorkbookPath, sourceWorkbookPath, WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
    assertTrue(OoxmlSecurityTestSupport.signatureValid(sourceWorkbookPath));

    Path copiedWorkbookPath = sourceWorkbookPath.getParent().resolve("copied-signed.xlsx");
    WorkbookFiles.copySourceWorkbook(
        sourceWorkbookPath, copiedWorkbookPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
    assertArrayEquals(
        Files.readAllBytes(sourceWorkbookPath), Files.readAllBytes(copiedWorkbookPath));

    Path deletedTempFile =
        ManagedTempFiles.createManagedTempFile("gridgrind-delete-if-exists-", ".tmp");
    WorkbookFiles.deleteIfExists(null);
    WorkbookFiles.deleteIfExists(deletedTempFile);
    assertFalse(Files.exists(deletedTempFile));

    Path nonEmptyDirectory =
        ManagedTempFiles.createManagedTempDirectory("gridgrind-delete-if-exists-dir-");
    Files.writeString(nonEmptyDirectory.resolve("child.txt"), "keep");
    WorkbookFiles.deleteIfExists(nonEmptyDirectory);
    assertTrue(Files.exists(nonEmptyDirectory));
    WorkbookFiles.deleteTreeIfExists(null);
    WorkbookFiles.deleteTreeIfExists(nonEmptyDirectory.resolve("missing-cleanup-root"));
    WorkbookFiles.deleteTreeIfExists(nonEmptyDirectory);
    assertFalse(Files.exists(nonEmptyDirectory));

    OoxmlEncryptionReport encryptedSnapshot =
        new OoxmlEncryptionReport.Encrypted(
            ExcelOoxmlEncryptionMode.AGILE,
            ExcelOoxmlCipherAlgorithm.AES_256,
            ExcelOoxmlHashAlgorithm.SHA_512,
            ExcelOoxmlChainingMode.CBC,
            256,
            16,
            100_000);
    IllegalArgumentException plaintextPreservationFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                SecurityAdapter.effectiveOptions(
                    OoxmlPackageSecurityReport.none(),
                    Optional.empty(),
                    new ResolvedPersistenceSecurity(
                        new OoxmlPersistenceEncryptionSpec.PreserveSource(),
                        new OoxmlPersistenceSignatureSpec.None())));
    assertTrue(plaintextPreservationFailure.getMessage().contains("requires an encrypted source"));

    IllegalStateException missingPasswordFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                SecurityAdapter.effectiveOptions(
                    new OoxmlPackageSecurityReport(encryptedSnapshot, java.util.List.of()),
                    Optional.empty(),
                    new ResolvedPersistenceSecurity(
                        new OoxmlPersistenceEncryptionSpec.PreserveSource(),
                        new OoxmlPersistenceSignatureSpec.None())));
    assertTrue(missingPasswordFailure.getMessage().contains("verified source password"));

    ResolvedPersistenceSecurity preservedOptions =
        SecurityAdapter.effectiveOptions(
            new OoxmlPackageSecurityReport(encryptedSnapshot, java.util.List.of()),
            Optional.of("persist-pass"),
            new ResolvedPersistenceSecurity(
                new OoxmlPersistenceEncryptionSpec.PreserveSource(),
                new OoxmlPersistenceSignatureSpec.None()));
    assertEquals(
        ExcelOoxmlWriteCipher.AES_256,
        assertInstanceOf(ResolvedPersistenceEncryption.Encrypt.class, preservedOptions.encryption())
            .options()
            .cipher());
    assertEquals(
        ExcelOoxmlWriteHash.SHA_512,
        assertInstanceOf(ResolvedPersistenceEncryption.Encrypt.class, preservedOptions.encryption())
            .options()
            .hash());

    OoxmlEncryptionReport standardSnapshot =
        new OoxmlEncryptionReport.Encrypted(
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
                SecurityAdapter.effectiveOptions(
                    new OoxmlPackageSecurityReport(standardSnapshot, java.util.List.of()),
                    Optional.of("persist-pass"),
                    new ResolvedPersistenceSecurity(
                        new OoxmlPersistenceEncryptionSpec.PreserveSource(),
                        new OoxmlPersistenceSignatureSpec.None())));
    assertTrue(standardPreservationFailure.getMessage().contains("not auto-preservable"));

    OoxmlEncryptionReport nonCbcSnapshot =
        new OoxmlEncryptionReport.Encrypted(
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
                SecurityAdapter.effectiveOptions(
                    new OoxmlPackageSecurityReport(nonCbcSnapshot, java.util.List.of()),
                    Optional.of("persist-pass"),
                    new ResolvedPersistenceSecurity(
                        new OoxmlPersistenceEncryptionSpec.PreserveSource(),
                        new OoxmlPersistenceSignatureSpec.None())));
    assertTrue(nonCbcPreservationFailure.getMessage().contains("chaining mode"));
    assertTrue(nonCbcPreservationFailure.getMessage().contains("CFB"));

    OoxmlEncryptionReport unsupportedCipherSnapshot =
        new OoxmlEncryptionReport.Encrypted(
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
                SecurityAdapter.effectiveOptions(
                    new OoxmlPackageSecurityReport(unsupportedCipherSnapshot, java.util.List.of()),
                    Optional.of("persist-pass"),
                    new ResolvedPersistenceSecurity(
                        new OoxmlPersistenceEncryptionSpec.PreserveSource(),
                        new OoxmlPersistenceSignatureSpec.None())));
    assertTrue(unsupportedCipherPreservationFailure.getMessage().contains("cipher AES_128"));

    OoxmlEncryptionReport unsupportedHashSnapshot =
        new OoxmlEncryptionReport.Encrypted(
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
                SecurityAdapter.effectiveOptions(
                    new OoxmlPackageSecurityReport(unsupportedHashSnapshot, java.util.List.of()),
                    Optional.of("persist-pass"),
                    new ResolvedPersistenceSecurity(
                        new OoxmlPersistenceEncryptionSpec.PreserveSource(),
                        new OoxmlPersistenceSignatureSpec.None())));
    assertTrue(unsupportedHashPreservationFailure.getMessage().contains("hash SHA_1"));
  }

  private static void assertSigningMaterialAndAliasBranches(
      OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook,
      KeyStore signingKeyStore,
      ResolvedSignature signatureOptions)
      throws IOException, GeneralSecurityException {
    assertNotNull(SecurityAdapter.signingMaterial(signatureOptions));

    ResolvedSignature wrongKeystorePassword =
        new ResolvedSignature(
            signedWorkbook.pkcs12Path(),
            "wrong-password",
            signedWorkbook.keyPassword(),
            java.util.Optional.of(signedWorkbook.alias()),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            java.util.Optional.empty());
    InvalidSigningConfigurationException wrongKeystorePasswordFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                SecurityAdapter.loadSigningKeyStore(
                    signedWorkbook.pkcs12Path(), wrongKeystorePassword));
    assertTrue(
        wrongKeystorePasswordFailure.getMessage().contains("Failed to load signing material"));

    InvalidSigningConfigurationException missingSigningMaterialFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                SecurityAdapter.signingMaterial(
                    new ResolvedSignature(
                        signedWorkbook.pkcs12Path().resolveSibling("missing-signing-material.p12"),
                        signedWorkbook.keystorePassword(),
                        signedWorkbook.keyPassword(),
                        java.util.Optional.of(signedWorkbook.alias()),
                        ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                        java.util.Optional.empty())));
    assertTrue(missingSigningMaterialFailure.getMessage().contains("does not exist"));

    assertEquals(
        signedWorkbook.alias(),
        SecurityAdapter.resolveAlias(signingKeyStore, signedWorkbook.alias(), signatureOptions));

    InvalidSigningConfigurationException missingAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () -> SecurityAdapter.resolveAlias(signingKeyStore, "missing-alias", signatureOptions));
    assertTrue(missingAliasFailure.getMessage().contains("does not exist"));

    KeyStore emptyKeyStore = KeyStore.getInstance("PKCS12");
    emptyKeyStore.load(null, null);
    InvalidSigningConfigurationException noKeyAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () -> SecurityAdapter.resolveAlias(emptyKeyStore, null, signatureOptions));
    assertTrue(noKeyAliasFailure.getMessage().contains("does not contain a private-key entry"));
  }

  private static void assertCertificateOnlyAndUninitializedKeystoreBranches(
      OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook,
      Certificate certificate,
      ResolvedSignature signatureOptions)
      throws IOException, GeneralSecurityException {
    KeyStore certificateOnlyKeyStore = KeyStore.getInstance("PKCS12");
    certificateOnlyKeyStore.load(null, null);
    certificateOnlyKeyStore.setCertificateEntry("certificate-only", certificate);

    InvalidSigningConfigurationException nonPrivateKeyFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                SecurityAdapter.signingPrivateKey(
                    certificateOnlyKeyStore,
                    signedWorkbook.pkcs12Path(),
                    "certificate-only",
                    signatureOptions));
    assertTrue(nonPrivateKeyFailure.getMessage().contains("does not resolve to a private key"));

    InvalidSigningConfigurationException certificateChainFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                SecurityAdapter.signingCertificateChain(
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
                SecurityAdapter.resolveSigningAlias(
                    uninitializedKeyStore, signedWorkbook.pkcs12Path(), signatureOptions));
    assertTrue(aliasInspectionFailure.getMessage().contains("Failed to inspect signing aliases"));

    InvalidSigningConfigurationException privateKeyLoadFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                SecurityAdapter.signingPrivateKey(
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
                SecurityAdapter.signingCertificateChain(
                    uninitializedKeyStore, signedWorkbook.pkcs12Path(), signedWorkbook.alias()));
    assertTrue(
        chainLoadFailure.getMessage().contains("Failed to load the signing certificate chain"));
  }

  private static void assertSyntheticKeystoreBranches(
      OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook,
      KeyStore signingKeyStore,
      Certificate certificate,
      ResolvedSignature signatureOptions)
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
            () -> SecurityAdapter.resolveAlias(multiAliasKeyStore, null, signatureOptions));
    assertTrue(multipleAliasesFailure.getMessage().contains("multiple private-key aliases"));

    KeyStore certificateOnlyAliasKeyStore = KeyStore.getInstance("PKCS12");
    certificateOnlyAliasKeyStore.load(null, null);
    certificateOnlyAliasKeyStore.setCertificateEntry("certificate-only", certificate);
    InvalidSigningConfigurationException certificateOnlyAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                SecurityAdapter.resolveAlias(certificateOnlyAliasKeyStore, null, signatureOptions));
    assertTrue(
        certificateOnlyAliasFailure.getMessage().contains("does not contain a private-key entry"));

    KeyStore nonPrivateKeyStore = fakeKeyStore(certificate.getPublicKey(), null);
    InvalidSigningConfigurationException nonPrivateAliasFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () -> SecurityAdapter.resolveAlias(nonPrivateKeyStore, null, signatureOptions));
    assertTrue(
        nonPrivateAliasFailure.getMessage().contains("does not contain a private-key entry"));
    InvalidSigningConfigurationException reflectedNonPrivateKeyFailure =
        assertThrows(
            InvalidSigningConfigurationException.class,
            () ->
                SecurityAdapter.signingPrivateKey(
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
                SecurityAdapter.signingCertificateChain(
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
                SecurityAdapter.signingCertificateChain(
                    emptyChainKeyStore, signedWorkbook.pkcs12Path(), "alias"));
    assertTrue(emptyChainFailure.getMessage().contains("X.509 certificate chain"));

    Path signableWorkbookPath =
        ManagedTempFiles.createManagedTempFile("gridgrind-sign-description-", ".xlsx");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Signed")
          .cells()
          .setCell("A1", new CellValue.Text("Signed workbook"));
      workbook
          .persistence()
          .save(
              signableWorkbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }
    SecurityAdapter.signWorkbook(
        signableWorkbookPath,
        new ResolvedSignature(
            signedWorkbook.pkcs12Path(),
            signedWorkbook.keystorePassword(),
            signedWorkbook.keyPassword(),
            java.util.Optional.of(signedWorkbook.alias()),
            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
            java.util.Optional.of("GridGrind test signature")));
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
    @SuppressWarnings("PMD.ReturnEmptyCollectionRatherThanNull")
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
