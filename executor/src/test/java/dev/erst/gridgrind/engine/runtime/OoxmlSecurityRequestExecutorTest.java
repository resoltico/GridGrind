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
  void corruptedAndNonWorkbookSourcesUseTheRequestFormatCode() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-unopenable-workbook-");
    Path nonZip = Files.writeString(directory.resolve("not-a-workbook.xlsx"), "not a zip");
    Path truncatedZip =
        Files.write(
            directory.resolve("truncated-workbook.xlsx"), new byte[] {'P', 'K', 3, 4, 20, 0, 0, 0});
    Path nonWorkbookZip = directory.resolve("not-a-workbook-package.xlsx");
    try (var zip = new java.util.zip.ZipOutputStream(Files.newOutputStream(nonWorkbookZip))) {
      zip.putNextEntry(new java.util.zip.ZipEntry("readme.txt"));
      zip.write("not an OOXML workbook".getBytes(java.nio.charset.StandardCharsets.UTF_8));
      zip.closeEntry();
    }

    for (Path malformedSource : List.of(nonZip, truncatedZip, nonWorkbookZip)) {
      assertUnopenableSource(malformedSource);
    }
  }

  private static void assertUnopenableSource(Path malformedSource) {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(malformedSource.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        inspect(
                            "workbook",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_OPENABLE, failure.problem().code());
    assertEquals(GridGrindProblemCategory.REQUEST, failure.problem().category());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
    assertTrue(failure.journal().steps().isEmpty());
  }

  @Test
  void existingSourceWriteRejectsOmittedTotalSecurityPolicy() throws IOException {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-signed-copy-"));
    Path copiedWorkbook = signedWorkbook.workbookPath().getParent().resolve("signed-copy.xlsx");

    WorkbookResult.Failure failure =
        failure(
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

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
    assertFalse(Files.exists(copiedWorkbook));
  }

  @Test
  void explicitNonePoliciesProduceAnUnsignedPlaintextExistingSourceOutput() throws IOException {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(
            Files.createTempDirectory("gridgrind-protocol-signed-unsigned-"));
    Path output = signedWorkbook.workbookPath().getParent().resolve("unsigned-output.xlsx");

    WorkbookResult.Success persisted =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(
                        signedWorkbook.workbookPath().toString()),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        output.toString(),
                        WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                        OoxmlPersistenceSecurityInput.none()),
                    mutations(
                        mutate(
                            new CellSelector.ByAddress("Signed", "B1"),
                            new CellMutationAction.SetCell(textCell("unsigned")))),
                    inspections())));
    assertEquals(output.toAbsolutePath().toString(), savedPath(persisted));

    WorkbookResult.Success reopened =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(output.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        inspect(
                            "security",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetPackageSecurity())))));
    WorkbookInspectionResult.PackageSecurityResult security =
        read(reopened, "security", WorkbookInspectionResult.PackageSecurityResult.class);
    assertInstanceOf(OoxmlEncryptionReport.None.class, security.security().encryption());
    assertTrue(security.security().signatures().isEmpty());
  }

  @Test
  void newSourceWritesDefaultToPlaintextAndUnsignedWhenSecurityIsOmitted() throws IOException {
    Path output =
        Files.createTempDirectory("gridgrind-protocol-new-security-default-").resolve("new.xlsx");

    WorkbookResult.Success persisted =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        output.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("New"),
                            new WorkbookMutationAction.EnsureSheet())),
                    inspections())));
    assertEquals(output.toAbsolutePath().toString(), savedPath(persisted));

    WorkbookResult.Success reopened =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(output.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    List.of(),
                    List.of(
                        inspect(
                            "security",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetPackageSecurity())))));
    WorkbookInspectionResult.PackageSecurityResult security =
        read(reopened, "security", WorkbookInspectionResult.PackageSecurityResult.class);
    assertInstanceOf(OoxmlEncryptionReport.None.class, security.security().encryption());
    assertTrue(security.security().signatures().isEmpty());
  }

  @Test
  void preserveSourceOnPlaintextExistingWorkbookFailsDuringPreflight() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-preserve-plaintext-");
    Path source = directory.resolve("plain.xlsx");
    try (org.apache.poi.xssf.usermodel.XSSFWorkbook workbook =
        new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      workbook.createSheet("Source").createRow(0).createCell(0).setCellValue("plain");
      try (java.io.OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(source.toString()),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        directory.resolve("output.xlsx").toString(),
                        WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                        new OoxmlPersistenceSecurityInput(
                            new OoxmlPersistenceEncryptionInput.PreserveSource(),
                            new OoxmlPersistenceSignatureInput.None())),
                    mutations(
                        mutate(
                            new SheetSelector.ByName("Source"),
                            new WorkbookMutationAction.EnsureSheet())))));

    assertEquals(GridGrindProblemCode.ENCRYPTION_SOURCE_NOT_ENCRYPTED, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
    assertTrue(failure.journal().steps().isEmpty());
    assertFalse(Files.exists(directory.resolve("output.xlsx")));
  }

  @Test
  void doctorBatchesIndependentPreservationAndSigningMaterialFailures() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-preservation-signing-doctor-");
    Path source = directory.resolve("plain.xlsx");
    try (org.apache.poi.xssf.usermodel.XSSFWorkbook workbook =
        new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
      workbook.createSheet("Source").createRow(0).createCell(0).setCellValue("plain");
      try (java.io.OutputStream output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    Path output = directory.resolve("output.xlsx");
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.ExistingFile(source.toString()),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                output.toString(),
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                new OoxmlPersistenceSecurityInput(
                    new OoxmlPersistenceEncryptionInput.PreserveSource(),
                    new OoxmlPersistenceSignatureInput.Sign(
                        new OoxmlSignatureInput(
                            "missing-signing-material.p12",
                            "keystore-pass",
                            "key-pass",
                            java.util.Optional.empty(),
                            dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm
                                .SHA256,
                            java.util.Optional.empty())))),
            mutations(
                mutate(
                    new SheetSelector.ByName("Source"), new WorkbookMutationAction.EnsureSheet())),
            inspections());

    RequestDoctorReport report =
        new GridGrindRequestDoctor()
            .diagnose(request, ExecutionInputBindingsFixtureSupport.bindings(directory));

    assertFalse(report.valid());
    assertEquals(
        List.of(
            GridGrindProblemCode.ENCRYPTION_SOURCE_NOT_ENCRYPTED,
            GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION),
        report.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
    assertFalse(Files.exists(output));
  }

  @Test
  void doctorBatchesWrongSourcePasswordAndMissingSigningMaterialBeforeExecution()
      throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook encryptedWorkbook =
        OoxmlSecurityTestSupport.createEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-password-signing-doctor-"));
    Path directory = encryptedWorkbook.workbookPath().getParent();
    Path output = directory.resolve("output.xlsx");
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.ExistingFile(
                encryptedWorkbook.workbookPath().toString(),
                new OoxmlOpenSecurityInput(java.util.Optional.of("wrong-password"))),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                output.toString(),
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                new OoxmlPersistenceSecurityInput(
                    new OoxmlPersistenceEncryptionInput.None(),
                    new OoxmlPersistenceSignatureInput.Sign(
                        new OoxmlSignatureInput(
                            "missing-signing-material.p12",
                            "keystore-pass",
                            "key-pass",
                            java.util.Optional.empty(),
                            dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm
                                .SHA256,
                            java.util.Optional.empty())))),
            mutations(
                mutate(
                    new CellSelector.ByAddress("Encrypted", "B1"),
                    new CellMutationAction.SetCell(textCell("not written")))),
            inspections());

    RequestDoctorReport report =
        new GridGrindRequestDoctor()
            .diagnose(request, ExecutionInputBindingsFixtureSupport.bindings(directory));
    WorkbookResult.Failure execution =
        failure(
            ExecutionContextFixtureSupport.execute(new DefaultGridGrindRequestExecutor(), request));

    assertFalse(report.valid());
    assertEquals(
        List.of(
            GridGrindProblemCode.INVALID_WORKBOOK_PASSWORD,
            GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION),
        report.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
    assertEquals(GridGrindProblemCode.INVALID_WORKBOOK_PASSWORD, execution.problem().code());
    assertTrue(execution.journal().steps().isEmpty());
    assertFalse(Files.exists(output));
  }

  @Test
  void preserveSourceRejectsAReadableButWriteIncompatibleEnvelopeBeforeMutation()
      throws IOException {
    OoxmlSecurityTestSupport.EncryptedWorkbook standardEncryptedWorkbook =
        OoxmlSecurityTestSupport.createLegacyStandardEncryptedWorkbook(
            Files.createTempDirectory("gridgrind-preserve-standard-"));
    Path output = standardEncryptedWorkbook.workbookPath().getParent().resolve("output.xlsx");

    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(
                        standardEncryptedWorkbook.workbookPath().toString(),
                        new OoxmlOpenSecurityInput(
                            java.util.Optional.of(standardEncryptedWorkbook.password()))),
                    new WorkbookPlan.WorkbookPersistence.SaveAs(
                        output.toString(),
                        WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                        new OoxmlPersistenceSecurityInput(
                            new OoxmlPersistenceEncryptionInput.PreserveSource(),
                            new OoxmlPersistenceSignatureInput.None())),
                    mutations(
                        mutate(
                            new CellSelector.ByAddress("Encrypted", "B1"),
                            new CellMutationAction.SetCell(textCell("changed")))),
                    inspections())));

    assertEquals(GridGrindProblemCode.ENCRYPTION_SOURCE_NOT_PRESERVABLE, failure.problem().code());
    assertEquals(GridGrindProblemCategory.SECURITY, failure.problem().category());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
    assertTrue(failure.journal().steps().isEmpty());
    assertFalse(Files.exists(output));
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
                            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput
                                .Encrypt(
                                new OoxmlEncryptionInput(
                                    OoxmlSecurityTestSupport.ENCRYPTION_PASSWORD,
                                    ExcelOoxmlWriteCipher.AES_192,
                                    ExcelOoxmlWriteHash.SHA_384)),
                            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.Sign(
                                new OoxmlSignatureInput(
                                    signingMaterial.pkcs12Path().toString(),
                                    signingMaterial.keystorePassword(),
                                    signingMaterial.keyPassword(),
                                    java.util.Optional.of(signingMaterial.alias()),
                                    dev.erst.gridgrind.excel.foundation
                                        .ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                                    java.util.Optional.of("GridGrind protocol signing test"))))),
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
                            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput
                                .None(),
                            new dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.Sign(
                                new OoxmlSignatureInput(
                                    outputPath.resolveSibling("missing.p12").toString(),
                                    "keystore-pass",
                                    "key-pass",
                                    java.util.Optional.empty(),
                                    dev.erst.gridgrind.excel.foundation
                                        .ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                                    java.util.Optional.empty())))),
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Secure"),
                            new WorkbookMutationAction.EnsureSheet())),
                    List.of())));

    assertEquals(GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION, failure.problem().code());
    assertEquals(GridGrindProblemCategory.SECURITY, failure.problem().category());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
    assertTrue(failure.journal().steps().isEmpty());
    assertFalse(Files.exists(outputPath));
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
