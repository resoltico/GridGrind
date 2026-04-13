package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.OoxmlSecurityTestSupport;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(
                            encryptedWorkbook.workbookPath().toString(),
                            new OoxmlOpenSecurityInput(encryptedWorkbook.password())),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(),
                        List.of(
                            new WorkbookReadOperation.GetPackageSecurity("security"),
                            new WorkbookReadOperation.GetCells(
                                "cells", "Encrypted", List.of("A1"))))));

    WorkbookReadResult.PackageSecurityResult security =
        read(success, "security", WorkbookReadResult.PackageSecurityResult.class);
    WorkbookReadResult.CellsResult cells =
        read(success, "cells", WorkbookReadResult.CellsResult.class);

    assertTrue(security.security().encryption().encrypted());
    assertEquals(List.of(), security.security().signatures());
    assertEquals(
        "Encrypted workbook",
        assertInstanceOf(GridGrindResponse.CellReport.TextReport.class, cells.cells().getFirst())
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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(
                            encryptedWorkbook.workbookPath().toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(),
                        List.of(new WorkbookReadOperation.GetWorkbookSummary("workbook")))));
    GridGrindResponse.Failure wrongPassword =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(
                            encryptedWorkbook.workbookPath().toString(),
                            new OoxmlOpenSecurityInput("wrong-password")),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(),
                        List.of(new WorkbookReadOperation.GetWorkbookSummary("workbook")))));

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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(
                            signedWorkbook.workbookPath().toString()),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(copiedWorkbook.toString()),
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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(
                            signedWorkbook.workbookPath().toString()),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(outputPath.toString()),
                        List.of(
                            new WorkbookOperation.SetCell(
                                "Signed", "C1", new CellInput.Text("Touch"))),
                        List.of())));

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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(
                            securedWorkbook.toString(),
                            new OoxmlPersistenceSecurityInput(
                                new OoxmlEncryptionInput(
                                    OoxmlSecurityTestSupport.ENCRYPTION_PASSWORD, null),
                                new OoxmlSignatureInput(
                                    signingMaterial.pkcs12Path().toString(),
                                    signingMaterial.keystorePassword(),
                                    signingMaterial.keyPassword(),
                                    signingMaterial.alias(),
                                    null,
                                    "GridGrind protocol signing test"))),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Secure"),
                            new WorkbookOperation.SetCell(
                                "Secure", "A1", new CellInput.Text("Secured"))),
                        List.of())));

    assertEquals(securedWorkbook.toAbsolutePath().toString(), savedPath(persisted));

    GridGrindResponse.Success reopened =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(
                            securedWorkbook.toString(),
                            new OoxmlOpenSecurityInput(
                                OoxmlSecurityTestSupport.ENCRYPTION_PASSWORD)),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(),
                        List.of(
                            new WorkbookReadOperation.GetPackageSecurity("security"),
                            new WorkbookReadOperation.GetCells(
                                "cells", "Secure", List.of("A1"))))));

    WorkbookReadResult.PackageSecurityResult security =
        read(reopened, "security", WorkbookReadResult.PackageSecurityResult.class);
    WorkbookReadResult.CellsResult cells =
        read(reopened, "cells", WorkbookReadResult.CellsResult.class);

    assertTrue(security.security().encryption().encrypted());
    assertEquals(1, security.security().signatures().size());
    assertEquals(
        dev.erst.gridgrind.excel.ExcelOoxmlSignatureState.VALID,
        security.security().signatures().getFirst().state());
    assertEquals(
        "Secured",
        assertInstanceOf(GridGrindResponse.CellReport.TextReport.class, cells.cells().getFirst())
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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(
                            outputPath.toString(),
                            new OoxmlPersistenceSecurityInput(
                                null,
                                new OoxmlSignatureInput(
                                    outputPath.resolveSibling("missing.p12").toString(),
                                    "keystore-pass",
                                    "key-pass",
                                    null,
                                    null,
                                    null))),
                        List.of(new WorkbookOperation.EnsureSheet("Secure")),
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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(new WorkbookReadOperation.GetPackageSecurity("security")))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("unsupported read type: GET_PACKAGE_SECURITY"));
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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(
                            tamperedWorkbook.toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(),
                        List.of(new WorkbookReadOperation.GetPackageSecurity("security")))));

    WorkbookReadResult.PackageSecurityResult security =
        read(success, "security", WorkbookReadResult.PackageSecurityResult.class);
    assertEquals(
        dev.erst.gridgrind.excel.ExcelOoxmlSignatureState.INVALID,
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
      case GridGrindResponse.PersistenceOutcome.SavedAs savedAs -> savedAs.executionPath();
      case GridGrindResponse.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponse.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("Expected the request to persist a workbook");
    };
  }

  private static <T extends WorkbookReadResult> T read(
      GridGrindResponse.Success success, String requestId, Class<T> type) {
    return type.cast(
        success.reads().stream()
            .filter(result -> result.requestId().equals(requestId))
            .findFirst()
            .orElseThrow(() -> new AssertionError("Missing read result " + requestId)));
  }
}
