package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** End-to-end tests for encrypted and signed OOXML package handling. */
class ExcelOoxmlPackageSecuritySupportTest {
  @Test
  void encryptedWorkbookOpenRequiresCorrectPasswordAndReportsEncryptionFacts() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-ooxml-encrypted-open-"));

    assertThrows(
        WorkbookPasswordRequiredException.class,
        () -> ExcelWorkbook.open(encryptedWorkbook.workbookPath()));
    assertThrows(
        InvalidWorkbookPasswordException.class,
        () ->
            ExcelWorkbook.open(
                encryptedWorkbook.workbookPath(),
                new ExcelOoxmlOpenOptions.Encrypted("wrong-password")));

    try (ExcelWorkbook workbook =
        ExcelWorkbook.open(
            encryptedWorkbook.workbookPath(),
            new ExcelOoxmlOpenOptions.Encrypted(encryptedWorkbook.password()))) {
      assertEquals(
          "Encrypted workbook",
          assertInstanceOf(
                  ExcelCellSnapshot.TextSnapshot.class,
                  workbook.sheet("Encrypted").snapshotCell("A1"))
              .stringValue());

      WorkbookReadResult.PackageSecurityResult securityResult =
          assertInstanceOf(
              WorkbookReadResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(workbook, new WorkbookReadCommand.GetPackageSecurity("security"))
                  .getFirst());
      assertTrue(securityResult.security().encryption().encrypted());
      assertEquals(ExcelOoxmlEncryptionMode.AGILE, securityResult.security().encryption().mode());
      assertEquals(List.of(), securityResult.security().signatures());
    }
  }

  @Test
  void encryptedSourcePreservesEncryptionAcrossUnchangedAndMutatedSaves() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-ooxml-encrypted-save-"));
    Path unchangedCopy =
        encryptedWorkbook.workbookPath().getParent().resolve("encrypted-unchanged-copy.xlsx");
    Path mutatedCopy =
        encryptedWorkbook.workbookPath().getParent().resolve("encrypted-mutated-copy.xlsx");

    try (ExcelWorkbook workbook =
        ExcelWorkbook.open(
            encryptedWorkbook.workbookPath(),
            new ExcelOoxmlOpenOptions.Encrypted(encryptedWorkbook.password()))) {
      workbook.save(unchangedCopy);
    }
    assertEquals(
        "Encrypted workbook",
        OoxmlSecurityTestSupport.decryptedStringCell(
            unchangedCopy, encryptedWorkbook.password(), "Encrypted", "A1"));

    try (ExcelWorkbook workbook =
        ExcelWorkbook.open(
            encryptedWorkbook.workbookPath(),
            new ExcelOoxmlOpenOptions.Encrypted(encryptedWorkbook.password()))) {
      new WorkbookCommandExecutor()
          .apply(
              workbook,
              new WorkbookCommand.SetCell("Encrypted", "B2", ExcelCellValue.text("Mutated")));
      workbook.save(mutatedCopy);
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
            Files.createTempDirectory("gridgrind-ooxml-signed-save-"));
    Path resignedOutput =
        signedWorkbook.workbookPath().getParent().resolve("signed-resigned-output.xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.open(signedWorkbook.workbookPath())) {
      WorkbookReadResult.PackageSecurityResult beforeMutation =
          assertInstanceOf(
              WorkbookReadResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(workbook, new WorkbookReadCommand.GetPackageSecurity("security"))
                  .getFirst());
      assertEquals(1, beforeMutation.security().signatures().size());
      assertEquals(
          ExcelOoxmlSignatureState.VALID,
          beforeMutation.security().signatures().getFirst().state());

      new WorkbookCommandExecutor()
          .apply(
              workbook, new WorkbookCommand.SetCell("Signed", "C1", ExcelCellValue.text("Touch")));

      WorkbookReadResult.PackageSecurityResult afterMutation =
          assertInstanceOf(
              WorkbookReadResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(workbook, new WorkbookReadCommand.GetPackageSecurity("security"))
                  .getFirst());
      assertEquals(
          ExcelOoxmlSignatureState.INVALIDATED_BY_MUTATION,
          afterMutation.security().signatures().getFirst().state());

      IllegalArgumentException unsignedSaveFailure =
          assertThrows(IllegalArgumentException.class, () -> workbook.save(resignedOutput));
      assertTrue(unsignedSaveFailure.getMessage().contains("persistence.security.signature"));

      workbook.save(
          resignedOutput,
          new ExcelOoxmlPersistenceOptions(
              null,
              new ExcelOoxmlSignatureOptions(
                  signedWorkbook.pkcs12Path(),
                  signedWorkbook.keystorePassword(),
                  signedWorkbook.keyPassword(),
                  signedWorkbook.alias(),
                  ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                  "GridGrind test signature")));
    }

    assertTrue(OoxmlSecurityTestSupport.signatureValid(resignedOutput));
    try (ExcelWorkbook reopened = ExcelWorkbook.open(resignedOutput)) {
      WorkbookReadResult.PackageSecurityResult resignedSecurity =
          assertInstanceOf(
              WorkbookReadResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(reopened, new WorkbookReadCommand.GetPackageSecurity("security"))
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
            Files.createTempDirectory("gridgrind-ooxml-signed-invalid-"));
    Path tamperedWorkbook =
        signedWorkbook.workbookPath().getParent().resolve("signed-invalid-tampered.xlsx");
    OoxmlSecurityTestSupport.tamperWorkbookCell(
        signedWorkbook.workbookPath(), tamperedWorkbook, "Signed", "B2", "Broken");

    assertFalse(OoxmlSecurityTestSupport.signatureValid(tamperedWorkbook));
    try (ExcelWorkbook workbook = ExcelWorkbook.open(tamperedWorkbook)) {
      WorkbookReadResult.PackageSecurityResult securityResult =
          assertInstanceOf(
              WorkbookReadResult.PackageSecurityResult.class,
              new WorkbookReadExecutor()
                  .apply(workbook, new WorkbookReadCommand.GetPackageSecurity("security"))
                  .getFirst());
      assertEquals(
          ExcelOoxmlSignatureState.INVALID,
          securityResult.security().signatures().getFirst().state());
    }
  }
}
