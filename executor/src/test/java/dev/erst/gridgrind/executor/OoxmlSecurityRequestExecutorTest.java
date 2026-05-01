package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.OoxmlSecurityTestSupport;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** End-to-end protocol tests for OOXML encryption and signing workflows. */
class OoxmlSecurityRequestExecutorTest {
  @Test
  void readsEncryptedWorkbookWithPasswordAndReportsPackageSecurity() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-encrypted-"));

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
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
                                new InspectionQuery.GetPackageSecurity()),
                            inspect(
                                "cells",
                                new CellSelector.ByAddresses("Encrypted", List.of("A1")),
                                new InspectionQuery.GetCells())))));

    InspectionResult.PackageSecurityResult security =
        read(success, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult cells = read(success, "cells", InspectionResult.CellsResult.class);

    assertTrue(security.security().encryption().encrypted());
    assertEquals(List.of(), security.security().signatures());
    assertEquals(
        "Encrypted workbook",
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
                cells.cells().getFirst())
            .stringValue());
  }

  @Test
  void encryptedWorkbookFailuresUseStableProblemCodes() throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-encrypted-failures-"));

    GridGrindResponse.Failure missingPassword =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(
                            encryptedWorkbook.workbookPath().toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        List.of(
                            inspect(
                                "workbook",
                                new WorkbookSelector.Current(),
                                new InspectionQuery.GetWorkbookSummary())))));
    GridGrindResponse.Failure wrongPassword =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
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
                                new InspectionQuery.GetWorkbookSummary())))));

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

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(
                            signedWorkbook.workbookPath().toString()),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(copiedWorkbook.toString()),
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

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(
                            signedWorkbook.workbookPath().toString()),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(outputPath.toString()),
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

    GridGrindResponse.Success persisted =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(
                            securedWorkbook.toString(),
                            new OoxmlPersistenceSecurityInput(
                                new OoxmlEncryptionInput(
                                    OoxmlSecurityTestSupport.ENCRYPTION_PASSWORD,
                                    ExcelOoxmlEncryptionMode.AGILE),
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

    GridGrindResponse.Success reopened =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(
                            securedWorkbook.toString(),
                            new OoxmlOpenSecurityInput(
                                java.util.Optional.of(
                                    OoxmlSecurityTestSupport.ENCRYPTION_PASSWORD))),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        List.of(
                            inspect(
                                "security",
                                new WorkbookSelector.Current(),
                                new InspectionQuery.GetPackageSecurity()),
                            inspect(
                                "cells",
                                new CellSelector.ByAddresses("Secure", List.of("A1")),
                                new InspectionQuery.GetCells())))));

    InspectionResult.PackageSecurityResult security =
        read(reopened, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult cells =
        read(reopened, "cells", InspectionResult.CellsResult.class);

    assertTrue(security.security().encryption().encrypted());
    assertEquals(1, security.security().signatures().size());
    assertEquals(
        dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState.VALID,
        security.security().signatures().getFirst().state());
    assertEquals(
        "Secured",
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.CellReport.TextReport.class,
                cells.cells().getFirst())
            .stringValue());
  }

  @Test
  void invalidSigningConfigurationUsesStableProblemCode() throws IOException {
    Path outputPath =
        Files.createTempDirectory("gridgrind-protocol-signing-invalid-").resolve("bad.xlsx");

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(
                            outputPath.toString(),
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
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.save(workbookPath);
    }

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            inspect(
                                "security",
                                new WorkbookSelector.Current(),
                                new InspectionQuery.GetPackageSecurity())))));

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

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(tamperedWorkbook.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        List.of(
                            inspect(
                                "security",
                                new WorkbookSelector.Current(),
                                new InspectionQuery.GetPackageSecurity())))));

    InspectionResult.PackageSecurityResult security =
        read(success, "security", InspectionResult.PackageSecurityResult.class);
    assertEquals(
        dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState.INVALID,
        security.security().signatures().getFirst().state());
  }

  private static GridGrindResponse.Success success(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Success.class, response);
  }

  private static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Failure.class, response);
  }

  private static String savedPath(GridGrindResponse.Success success) {
    return switch (success.persistence()) {
      case GridGrindResponsePersistence.PersistenceOutcome.SavedAs savedAs ->
          savedAs.executionPath();
      case GridGrindResponsePersistence.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponsePersistence.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("Expected the request to persist a workbook");
    };
  }

  private static <T extends InspectionResult> T read(
      GridGrindResponse.Success success, String stepId, Class<T> type) {
    return inspection(success, stepId, type);
  }
}
