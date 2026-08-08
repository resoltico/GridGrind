package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.OoxmlSecurityTestSupport;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteHash;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** End-to-end protocol tests for OOXML encryption and signing workflows. */
class OoxmlSecurityRequestExecutorTest extends DefaultGridGrindRequestExecutorTestSupport {
  @Test
  void readsEncryptedWorkbookWithPasswordAndReportsPackageSecurity() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-encrypted-"));

    WorkbookResult.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(
                        encryptedWorkbook.workbookPath().toString(),
                        new OoxmlOpenSecurityInput(
                            java.util.Optional.of(encryptedWorkbook.password()))),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        inspect(
                            "security",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetPackageSecurity()),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Encrypted", List.of("A1")),
                            new SheetIntrospectionQuery.GetCells())))));

    WorkbookInspectionResult.PackageSecurityResult security =
        read(success, "security", WorkbookInspectionResult.PackageSecurityResult.class);
    SheetInspectionResult.CellsResult cells =
        read(success, "cells", SheetInspectionResult.CellsResult.class);

    assertInstanceOf(OoxmlEncryptionReport.Encrypted.class, security.security().encryption());
    assertEquals(List.of(), security.security().signatures());
    assertEquals(
        "Encrypted workbook",
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
                cells.cells().getFirst())
            .textValue()
            .orElseThrow());
  }

  @Test
  void encryptedWorkbookFailuresUseStableProblemCodes() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-encrypted-failures-"));

    WorkbookResult.Failure missingPassword =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(
                        encryptedWorkbook.workbookPath().toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        inspect(
                            "workbook",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookSummary())))));
    WorkbookResult.Failure wrongPassword =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(
                        encryptedWorkbook.workbookPath().toString(),
                        new OoxmlOpenSecurityInput(java.util.Optional.of("wrong-password"))),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        inspect(
                            "workbook",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.WORKBOOK_PASSWORD_REQUIRED, missingPassword.problem().code());
    assertEquals(GridGrindProblemCategory.SECURITY, missingPassword.problem().category());
    assertEquals(GridGrindProblemCode.INVALID_WORKBOOK_PASSWORD, wrongPassword.problem().code());
    assertEquals(GridGrindProblemCategory.SECURITY, wrongPassword.problem().category());
  }

  @Test
  void unchangedSignedSourceCanBeSavedAsWithoutDroppingItsSignature() throws IOException {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-signed-copy-"));
    Path copiedWorkbook = signedWorkbook.workbookPath().getParent().resolve("signed-copy.xlsx");

    WorkbookResult.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(
                        signedWorkbook.workbookPath().toString()),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        copiedWorkbook.toString(),
                        WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                    List.of(),
                    List.of())));

    assertEquals(copiedWorkbook.toAbsolutePath().toString(), savedPath(success));
    assertTrue(OoxmlSecurityTestSupport.signatureValid(copiedWorkbook));
  }

  @Test
  void mutatedSignedSourceRequiresExplicitResigning() throws IOException {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-signed-mutated-"));
    Path outputPath = signedWorkbook.workbookPath().getParent().resolve("signed-mutated.xlsx");

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(
                        signedWorkbook.workbookPath().toString()),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        outputPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                    mutations(
                        mutate(
                            new CellSelector.ByAddress("Signed", "C1"),
                            new CellMutationAction.SetCell(textCell("Touch")))),
                    inspections())));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("persistence.security.signature"));
  }

  @Test
  void saveAsCanEncryptAndSignNewWorkbookThenReadBackBothFacts() throws IOException {
    OoxmlSecurityTestSupport.SignedWorkbook signingMaterial =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-signing-material-"));
    Path securedWorkbook =
        signingMaterial.workbookPath().getParent().resolve("secured-output.xlsx");

    WorkbookResult.Success persisted =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        securedWorkbook.toString(),
                        WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                        new OoxmlPersistenceSecurityInput(
                            new OoxmlEncryptionInput(
                                OoxmlSecurityTestSupport.ENCRYPTION_PASSWORD,
                                ExcelOoxmlWriteCipher.AES_192,
                                ExcelOoxmlWriteHash.SHA_384),
                            new OoxmlSignatureInput(
                                signingMaterial.pkcs12Path().toString(),
                                signingMaterial.keystorePassword(),
                                signingMaterial.keyPassword(),
                                java.util.Optional.of(signingMaterial.alias()),
                                dev.erst.gridgrind.excel.foundation
                                    .ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                                java.util.Optional.of("GridGrind protocol signing test")))),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Secure"),
                            new WorkbookMutationAction.EnsureSheet()),
                        mutate(
                            new CellSelector.ByAddress("Secure", "A1"),
                            new CellMutationAction.SetCell(textCell("Secured")))),
                    inspections())));

    assertEquals(securedWorkbook.toAbsolutePath().toString(), savedPath(persisted));

    WorkbookResult.Success reopened =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(
                        securedWorkbook.toString(),
                        new OoxmlOpenSecurityInput(
                            java.util.Optional.of(OoxmlSecurityTestSupport.ENCRYPTION_PASSWORD))),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        inspect(
                            "security",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetPackageSecurity()),
                        inspect(
                            "cells",
                            new CellSelector.ByAddresses("Secure", List.of("A1")),
                            new SheetIntrospectionQuery.GetCells())))));

    WorkbookInspectionResult.PackageSecurityResult security =
        read(reopened, "security", WorkbookInspectionResult.PackageSecurityResult.class);
    SheetInspectionResult.CellsResult cells =
        read(reopened, "cells", SheetInspectionResult.CellsResult.class);

    OoxmlEncryptionReport.Encrypted encryption =
        assertInstanceOf(OoxmlEncryptionReport.Encrypted.class, security.security().encryption());
    assertEquals(ExcelOoxmlEncryptionMode.AGILE, encryption.mode());
    assertEquals(ExcelOoxmlCipherAlgorithm.AES_192, encryption.cipherAlgorithm());
    assertEquals(ExcelOoxmlHashAlgorithm.SHA_384, encryption.hashAlgorithm());
    assertEquals(1, security.security().signatures().size());
    assertEquals(
        dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState.VALID,
        security.security().signatures().getFirst().state());
    assertTrue(
        security
            .security()
            .signatures()
            .getFirst()
            .signer()
            .orElseThrow()
            .subject()
            .contains("GridGrind Signing Test"));
    assertEquals(
        "Secured",
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
                cells.cells().getFirst())
            .textValue()
            .orElseThrow());
  }

  @Test
  void invalidSigningConfigurationUsesStableProblemCode() throws IOException {
    Path outputPath =
        Files.createTempDirectory("gridgrind-protocol-signing-invalid-").resolve("bad.xlsx");

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        outputPath.toString(),
                        WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                        new OoxmlPersistenceSecurityInput(
                            null,
                            new OoxmlSignatureInput(
                                outputPath.resolveSibling("missing.p12").toString(),
                                "keystore-pass",
                                "key-pass",
                                java.util.Optional.empty(),
                                dev.erst.gridgrind.excel.foundation
                                    .ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                                java.util.Optional.empty()))),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Secure"),
                            new WorkbookMutationAction.EnsureSheet())),
                    List.of())));

    assertEquals(GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION, failure.problem().code());
    assertEquals(GridGrindProblemCategory.SECURITY, failure.problem().category());
  }

  @Test
  void eventReadModeRejectsPackageSecurityReadsUpFront() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-protocol-event-security-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Ops");
      ExecutionContextFixtureSupport.saveWorkbook(workbook, workbookPath);
    }

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionModeInput.eventRead(),
                    null,
                    List.of(),
                    List.of(
                        inspect(
                            "security",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetPackageSecurity())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("GET_PACKAGE_SECURITY"));
    assertTrue(
        failure
            .problem()
            .message()
            .contains(
                dev.erst.gridgrind.contract.catalog.GridGrindContractText
                    .eventReadInspectionQueryTypePhrase()));
  }

  @Test
  void tamperedSignedWorkbookReadsBackInvalidSignatureState() throws IOException {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-signed-invalid-"));
    Path tamperedWorkbook =
        signedWorkbook.workbookPath().getParent().resolve("signed-invalid.xlsx");
    OoxmlSecurityTestSupport.tamperWorkbookCell(
        signedWorkbook.workbookPath(), tamperedWorkbook, "Signed", "B2", "Broken");

    WorkbookResult.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(tamperedWorkbook.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        inspect(
                            "security",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetPackageSecurity())))));

    WorkbookInspectionResult.PackageSecurityResult security =
        read(success, "security", WorkbookInspectionResult.PackageSecurityResult.class);
    assertEquals(
        dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState.INVALID,
        security.security().signatures().getFirst().state());
  }
}
