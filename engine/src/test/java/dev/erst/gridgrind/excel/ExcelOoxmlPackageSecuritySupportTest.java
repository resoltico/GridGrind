package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec;
import dev.erst.gridgrind.core.model.security.ResolvedOpenSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSecurity;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSignature;
import dev.erst.gridgrind.core.model.security.ResolvedSignature;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageInspectionSupport;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

/** End-to-end tests for encrypted and signed OOXML package handling. */
class ExcelOoxmlPackageSecuritySupportTest {
  @Test
  void openedWorkbooksRetainTheirSourcePathForPersistenceMetadata() throws IOException {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-source-path-"));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            signedWorkbook.workbookPath(), ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals(
          Optional.of(signedWorkbook.workbookPath().toAbsolutePath().normalize()),
          workbook.persistence().sourcePath());
    }
  }

  @Test
  void encryptedWorkbookOpenRequiresCorrectPasswordAndReportsEncryptionFacts() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-encrypted-open-"));

    assertThrows(
        WorkbookPasswordRequiredException.class,
        () ->
            ExcelWorkbooks.open(
                encryptedWorkbook.workbookPath(),
                ExcelTempFileFactoryTestSupport.tempFileFactory()));
    assertThrows(
        InvalidWorkbookPasswordException.class,
        () ->
            ExcelWorkbooks.open(
                encryptedWorkbook.workbookPath(),
                new ResolvedOpenSecurity.Encrypted("wrong-password"),
                ExcelTempFileFactoryTestSupport.tempFileFactory()));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            encryptedWorkbook.workbookPath(),
            new ResolvedOpenSecurity.Encrypted(encryptedWorkbook.password()),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals(
          "Encrypted workbook",
          assertInstanceOf(
                  CellReport.TextReport.class,
                  workbook.sheet("Encrypted").cells().snapshotCell("A1"))
              .value()
              .orElseThrow()
              .value());

      dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult securityResult =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("security"),
                          new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                          new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                              .GetPackageSecurity()))
                  .getFirst());
      OoxmlEncryptionReport.Encrypted encryption =
          assertInstanceOf(
              OoxmlEncryptionReport.Encrypted.class, securityResult.security().encryption());
      assertEquals(ExcelOoxmlEncryptionMode.AGILE, encryption.mode());
      assertEquals(ExcelOoxmlCipherAlgorithm.AES_256, encryption.cipherAlgorithm());
      assertEquals(ExcelOoxmlHashAlgorithm.SHA_512, encryption.hashAlgorithm());
      assertEquals(List.of(), securityResult.security().signatures());
    }
  }

  @ParameterizedTest(name = "{index}: {0}/{1}")
  @MethodSource("supportedWriteEncryptionEnvelopes")
  void authoredEncryptedWorkbookReportsEverySupportedWriteCipherAndHash(
      ExcelOoxmlWriteCipher cipher, ExcelOoxmlWriteHash hash) throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-encrypted-envelope-"),
            cipher,
            hash);

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            encryptedWorkbook.workbookPath(),
            new ResolvedOpenSecurity.Encrypted(encryptedWorkbook.password()),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult securityResult =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("security"),
                          new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                          new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                              .GetPackageSecurity()))
                  .getFirst());
      OoxmlEncryptionReport.Encrypted encryption =
          assertInstanceOf(
              OoxmlEncryptionReport.Encrypted.class, securityResult.security().encryption());
      assertEquals(ExcelOoxmlEncryptionMode.AGILE, encryption.mode());
      assertEquals(expectedCipher(cipher), encryption.cipherAlgorithm());
      assertEquals(expectedHash(hash), encryption.hashAlgorithm());
    }
  }

  @Test
  void legacyStandardEncryptedWorkbookStillReadsAndReportsStandard() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createLegacyStandardEncryptedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-standard-open-"));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            encryptedWorkbook.workbookPath(),
            new ResolvedOpenSecurity.Encrypted(encryptedWorkbook.password()),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult securityResult =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("security"),
                          new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                          new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                              .GetPackageSecurity()))
                  .getFirst());
      OoxmlEncryptionReport.Encrypted encryption =
          assertInstanceOf(
              OoxmlEncryptionReport.Encrypted.class, securityResult.security().encryption());
      assertEquals(ExcelOoxmlEncryptionMode.STANDARD, encryption.mode());
    }
  }

  @Test
  void encryptedSourcePreservesEncryptionAcrossUnchangedAndMutatedSaves() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-encrypted-save-"));
    Path unchangedCopy =
        encryptedWorkbook.workbookPath().getParent().resolve("encrypted-unchanged-copy.xlsx");
    Path mutatedCopy =
        encryptedWorkbook.workbookPath().getParent().resolve("encrypted-mutated-copy.xlsx");

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            encryptedWorkbook.workbookPath(),
            new ResolvedOpenSecurity.Encrypted(encryptedWorkbook.password()),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      workbook
          .persistence()
          .save(
              unchangedCopy,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              new ResolvedPersistenceSecurity(
                  new OoxmlPersistenceEncryptionSpec.PreserveSource(),
                  new OoxmlPersistenceSignatureSpec.None()),
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }
    assertEquals(
        "Encrypted workbook",
        OoxmlSecurityTestSupport.decryptedStringCell(
            unchangedCopy, encryptedWorkbook.password(), "Encrypted", "A1"));

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            encryptedWorkbook.workbookPath(),
            new ResolvedOpenSecurity.Encrypted(encryptedWorkbook.password()),
            ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      new WorkbookExecutionEngine()
          .apply(
              workbook,
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Encrypted"), new CellAddress("B2")),
                  new ResolvedMutationAction.SetCell(new CellValue.Text("Mutated"))));
      workbook
          .persistence()
          .save(
              mutatedCopy,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              new ResolvedPersistenceSecurity(
                  new OoxmlPersistenceEncryptionSpec.PreserveSource(),
                  new OoxmlPersistenceSignatureSpec.None()),
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertEquals(
        "Mutated",
        OoxmlSecurityTestSupport.decryptedStringCell(
            mutatedCopy, encryptedWorkbook.password(), "Encrypted", "B2"));
  }

  @Test
  void signedWorkbookReportsValidInvalidatedAndResignedStates() throws IOException {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-signed-save-"));
    Path unsignedOutput =
        signedWorkbook.workbookPath().getParent().resolve("signed-unsigned-output.xlsx");
    Path resignedOutput =
        signedWorkbook.workbookPath().getParent().resolve("signed-resigned-output.xlsx");

    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            signedWorkbook.workbookPath(), ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult beforeMutation =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("security"),
                          new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                          new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                              .GetPackageSecurity()))
                  .getFirst());
      assertEquals(1, beforeMutation.security().signatures().size());
      assertEquals(
          ExcelOoxmlSignatureState.VALID,
          beforeMutation.security().signatures().getFirst().state());

      new WorkbookExecutionEngine()
          .apply(
              workbook,
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Signed"), new CellAddress("C1")),
                  new ResolvedMutationAction.SetCell(new CellValue.Text("Touch"))));

      dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult afterMutation =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("security"),
                          new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                          new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                              .GetPackageSecurity()))
                  .getFirst());
      assertEquals(
          ExcelOoxmlSignatureState.INVALIDATED_BY_MUTATION,
          afterMutation.security().signatures().getFirst().state());

      workbook
          .persistence()
          .save(
              unsignedOutput,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ResolvedPersistenceSecurity.none(),
              ExcelTempFileFactoryTestSupport.tempFileFactory());

      workbook
          .persistence()
          .save(
              resignedOutput,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              new ResolvedPersistenceSecurity(
                  new OoxmlPersistenceEncryptionSpec.None(),
                  new ResolvedPersistenceSignature.Sign(
                      new ResolvedSignature(
                          signedWorkbook.pkcs12Path(),
                          signedWorkbook.keystorePassword(),
                          signedWorkbook.keyPassword(),
                          java.util.Optional.of(signedWorkbook.alias()),
                          ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                          java.util.Optional.of("GridGrind test signature")))),
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    assertTrue(
        ExcelOoxmlPackageInspectionSupport.inspectPackageSecurity(
                unsignedOutput, new OoxmlEncryptionReport.None())
            .signatures()
            .isEmpty());
    assertTrue(OoxmlSecurityTestSupport.signatureValid(resignedOutput));
    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(resignedOutput, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult resignedSecurity =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(
                      reopened,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("security"),
                          new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                          new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                              .GetPackageSecurity()))
                  .getFirst());
      assertEquals(
          ExcelOoxmlSignatureState.VALID,
          resignedSecurity.security().signatures().getFirst().state());
    }
  }

  @Test
  void tamperedSignedWorkbookReadsBackAsInvalidInsteadOfFailingOpen() throws IOException {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            ExcelTempFiles.createManagedTempDirectory("gridgrind-ooxml-signed-invalid-"));
    Path tamperedWorkbook =
        signedWorkbook.workbookPath().getParent().resolve("signed-invalid-tampered.xlsx");
    OoxmlSecurityTestSupport.tamperWorkbookCell(
        signedWorkbook.workbookPath(), tamperedWorkbook, "Signed", "B2", "Broken");

    assertFalse(OoxmlSecurityTestSupport.signatureValid(tamperedWorkbook));
    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(tamperedWorkbook, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult securityResult =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookInspectionResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("security"),
                          new dev.erst.gridgrind.core.plan.WorkbookSelector.Current(),
                          new dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery
                              .GetPackageSecurity()))
                  .getFirst());
      assertEquals(
          ExcelOoxmlSignatureState.INVALID,
          securityResult.security().signatures().getFirst().state());
    }
  }

  private static java.util.stream.Stream<Arguments> supportedWriteEncryptionEnvelopes() {
    return java.util.stream.Stream.of(
        Arguments.of(ExcelOoxmlWriteCipher.AES_256, ExcelOoxmlWriteHash.SHA_512),
        Arguments.of(ExcelOoxmlWriteCipher.AES_256, ExcelOoxmlWriteHash.SHA_384),
        Arguments.of(ExcelOoxmlWriteCipher.AES_256, ExcelOoxmlWriteHash.SHA_256),
        Arguments.of(ExcelOoxmlWriteCipher.AES_192, ExcelOoxmlWriteHash.SHA_512),
        Arguments.of(ExcelOoxmlWriteCipher.AES_192, ExcelOoxmlWriteHash.SHA_384),
        Arguments.of(ExcelOoxmlWriteCipher.AES_192, ExcelOoxmlWriteHash.SHA_256));
  }

  private static ExcelOoxmlCipherAlgorithm expectedCipher(ExcelOoxmlWriteCipher cipher) {
    return switch (cipher) {
      case AES_256 -> ExcelOoxmlCipherAlgorithm.AES_256;
      case AES_192 -> ExcelOoxmlCipherAlgorithm.AES_192;
    };
  }

  private static ExcelOoxmlHashAlgorithm expectedHash(ExcelOoxmlWriteHash hash) {
    return switch (hash) {
      case SHA_512 -> ExcelOoxmlHashAlgorithm.SHA_512;
      case SHA_384 -> ExcelOoxmlHashAlgorithm.SHA_384;
      case SHA_256 -> ExcelOoxmlHashAlgorithm.SHA_256;
    };
  }
}
