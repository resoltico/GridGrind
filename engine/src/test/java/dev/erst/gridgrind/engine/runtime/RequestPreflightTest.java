package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec;
import dev.erst.gridgrind.core.model.security.OoxmlSignatureSpec;
import dev.erst.gridgrind.core.model.value.CellGridSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.InvalidSigningConfigurationException;
import dev.erst.gridgrind.excel.OoxmlSecurityTestSupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceSignature;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSignatureOptions;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies phase-four source checks collect independent facts without workbook mutation. */
class RequestPreflightTest {
  @TempDir Path temporaryDirectory;

  @Test
  void preflightsSigningMaterialBeforeWorkbookExecution() throws Exception {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "signed.xlsx",
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                new OoxmlPersistenceSecuritySpec(
                    new dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec
                        .None(),
                    new OoxmlPersistenceSignatureSpec.Sign(
                        OoxmlSignatureSpec.sameKeyPassword(
                            "missing-signing-material.p12",
                            new dev.erst.gridgrind.core.source.SecretReference("signing-password"),
                            Optional.empty(),
                            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                            Optional.empty())))),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());

    RequestPreflight.Result result =
        RequestPreflight.verify(
            request,
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(request, temporaryDirectory)));
    try {
      assertEquals(
          List.of(GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION),
          result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
      assertTrue(result.preparedRequest().isEmpty());
    } finally {
      result.release();
    }

    assertThrows(
        InvalidSigningConfigurationException.class,
        () ->
            RequestPreflightPaths.verifySigningMaterial(
                new ExcelOoxmlPersistenceSignature.Sign(
                    new ExcelOoxmlSignatureOptions(
                        temporaryDirectory.resolve("missing-signing-material.p12"),
                        "password",
                        "password",
                        null,
                        ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                        null))));
    RequestPreflightPaths.verifySigningMaterial(new ExcelOoxmlPersistenceSignature.Unsigned());
  }

  @Test
  void collectsIndependentAssetAndWorkbookSourceFailuresWithoutAResolvedPlan() throws Exception {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("missing-source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("set-title"),
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetCell(
                        new CellValueSpec.Text(new TextSourceSpec.Utf8File("missing-title.txt")))),
                new MutationStep(
                    new StepId("set-subtitle"),
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A2")),
                    new CellMutationAction.SetCell(
                        new CellValueSpec.Text(
                            new TextSourceSpec.Utf8File("missing-subtitle.txt"))))));
    Path tempRoot = temporaryDirectory.resolve("scratch");

    RequestPreflight.Result result =
        RequestPreflight.verify(
            request,
            new ExecutionInputBindings(
                temporaryDirectory,
                tempRoot,
                ExecutionGrantTestSupport.forPlan(request, temporaryDirectory)));

    assertFalse(result.preparedRequest().isPresent());
    assertEquals(
        List.of(
            GridGrindProblemCode.WORKBOOK_NOT_FOUND,
            GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND,
            GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND),
        result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
    assertInstanceOf(ProblemContext.OpenWorkbook.class, result.problems().getFirst().context());
    assertInstanceOf(ProblemContext.ResolveInputs.class, result.problems().get(1).context());
    assertInstanceOf(ProblemContext.ResolveInputs.class, result.problems().get(2).context());
    assertFalse(Files.exists(temporaryDirectory.resolve("missing-source.xlsx")));
    assertFalse(Files.exists(tempRoot));
    assertTrue(
        result.problems().stream()
            .allMatch(
                problem ->
                    problem.context().type().endsWith("INPUTS")
                        || "OPEN_WORKBOOK".equals(problem.context().type())));
  }

  @Test
  void preflightsBoundSigningMaterialAfterItsPathHasBeenRead() throws Exception {
    OoxmlSecurityTestSupport.SignedWorkbook signedWorkbook =
        OoxmlSecurityTestSupport.createSignedWorkbook(temporaryDirectory.resolve("signing"));
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "signed.xlsx",
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                new OoxmlPersistenceSecuritySpec(
                    new dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec
                        .None(),
                    new OoxmlPersistenceSignatureSpec.Sign(
                        new OoxmlSignatureSpec(
                            signedWorkbook.pkcs12Path().toString(),
                            new dev.erst.gridgrind.core.source.SecretReference(
                                "wrong-keystore-password"),
                            Optional.of(
                                new dev.erst.gridgrind.core.source.SecretReference("key-password")),
                            Optional.of(signedWorkbook.alias()),
                            ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                            Optional.empty())))),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    ExecutionInputBindings rawBindings =
        new ExecutionInputBindings(
            temporaryDirectory,
            temporaryDirectory.resolve("scratch"),
            ExecutionGrantTestSupport.forPlan(
                request, temporaryDirectory, signedWorkbook.pkcs12Path()),
            reference -> "wrong-keystore-password".toCharArray());
    List<GridGrindProblemDetail.Problem> problems = new ArrayList<>();

    try (RequestPathAccess pathAccess =
        new RequestPathAccess(
            rawBindings.workingDirectory(),
            rawBindings.tempFileFactory(),
            rawBindings.executionGrant())) {
      RequestPreflightPaths.preflightPersistenceMaterial(
          request, rawBindings.withRequestPathAccess(pathAccess), problems);
      assertEquals(
          List.of(GridGrindProblemCode.INVALID_SIGNING_CONFIGURATION),
          problems.stream().map(GridGrindProblemDetail.Problem::code).toList());
    }
  }

  @Test
  void doctorBatchesIndependentPreflightFailuresAlongsideStaticFindings() {
    var analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": 1,
              "source": { "type": "EXISTING", "path": "missing-source.xlsx" },
              "persistence": { "type": "NONE" },
              "steps": [
                {
                  "stepId": "set-title",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "action": {
                    "type": "SET_CELL",
                    "value": {
                      "type": "TEXT",
                      "value": { "type": "UTF8_FILE", "path": "missing-title.txt" }
                    }
                  }
                }
              ]
            }
            """
                .getBytes(java.nio.charset.StandardCharsets.UTF_8));

    var report =
        new GridGrindRequestDoctor()
            .diagnose(
                analysis,
                ProblemContextRequestSurfaces.RequestInput.standardInput(),
                new ExecutionInputBindings(
                    temporaryDirectory,
                    temporaryDirectory.resolve("scratch"),
                    ExecutionGrantTestSupport.forPlan(
                        analysis.requireCompletePlan(), temporaryDirectory)));

    assertFalse(report.valid());
    assertEquals(
        List.of(
            GridGrindProblemCode.INVALID_REQUEST,
            GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND,
            GridGrindProblemCode.WORKBOOK_NOT_FOUND),
        report.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
    assertFalse(Files.exists(temporaryDirectory.resolve("missing-source.xlsx")));
    assertFalse(Files.exists(temporaryDirectory.resolve("scratch")));
  }

  @Test
  void retainsExactAuthoredSourcePathsForEveryBatchedInputFailure() {
    var analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": 1,
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [
                {
                  "stepId": "set-values",
                  "target": { "type": "RANGE_BY_RANGE", "sheetName": "Budget", "range": "A1:B1" },
                  "action": {
                    "type": "SET_RANGE",
                    "rows": {
                      "type": "TYPED",
                      "cells": [[
                        { "type": "TEXT", "value": { "type": "UTF8_FILE", "path": "missing-first.txt" } },
                        { "type": "TEXT", "value": { "type": "UTF8_FILE", "path": "missing-second.txt" } }
                      ]]
                    }
                  }
                }
              ]
            }
            """
                .getBytes(java.nio.charset.StandardCharsets.UTF_8));

    RequestPreflight.Result result =
        RequestPreflight.verify(
            analysis.requireCompletePlan(),
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(
                    analysis.requireCompletePlan(), temporaryDirectory)),
            analysis);

    assertEquals(2, result.problems().size());
    assertEquals(
        List.of(
            "steps[0].action.rows.cells[0][0].value.path",
            "steps[0].action.rows.cells[0][1].value.path"),
        result.problems().stream()
            .map(GridGrindProblemDetail.Problem::context)
            .map(ProblemContext.ResolveInputs.class::cast)
            .map(ProblemContext.ResolveInputs::json)
            .map(java.util.Optional::orElseThrow)
            .map(location -> location.jsonPathValue().orElseThrow())
            .toList());
    assertTrue(
        result.problems().stream()
            .map(GridGrindProblemDetail.Problem::context)
            .map(ProblemContext.ResolveInputs.class::cast)
            .map(ProblemContext.ResolveInputs::json)
            .map(java.util.Optional::orElseThrow)
            .allMatch(location -> location.byteOffsetValue().isPresent()));
  }

  @Test
  void collectsEveryIndependentSourceFailureWithinOneBoundStep() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("set-values"),
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetRange(
                        new CellGridSpec.Typed(
                            List.of(
                                List.of(
                                    new CellValueSpec.Text(
                                        new TextSourceSpec.Utf8File("missing-first.txt")),
                                    new CellValueSpec.Text(
                                        new TextSourceSpec.Utf8File("missing-second.txt")))))))));

    RequestPreflight.Result result =
        RequestPreflight.verify(
            request,
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(request, temporaryDirectory)));

    assertFalse(result.preparedRequest().isPresent());
    assertEquals(
        List.of(
            GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND,
            GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND),
        result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
    assertEquals(
        List.of("missing-first.txt", "missing-second.txt"),
        result.problems().stream()
            .map(GridGrindProblemDetail.Problem::context)
            .map(ProblemContext.ResolveInputs.class::cast)
            .map(ProblemContext.ResolveInputs::input)
            .map(ProblemContextWorkbookSurfaces.InputReference::inputPathValue)
            .map(java.util.Optional::orElseThrow)
            .toList());
  }

  @Test
  void collectsEveryRichTextLeafFailureWithinOneBoundStep() throws Exception {
    Path emptyRun = Files.createFile(temporaryDirectory.resolve("empty-run.txt"));
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("set-rich-text"),
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetRange(
                        new CellGridSpec.Typed(
                            List.of(
                                List.of(
                                    new CellValueSpec.RichText(
                                        List.of(
                                            new dev.erst.gridgrind.core.model.style.RichTextRunSpec(
                                                new TextSourceSpec.Utf8File("missing-run.txt"),
                                                java.util.Optional.empty()))),
                                    new CellValueSpec.RichText(
                                        List.of(
                                            new dev.erst.gridgrind.core.model.style.RichTextRunSpec(
                                                new TextSourceSpec.Utf8File("empty-run.txt"),
                                                java.util.Optional.empty()))))))))));

    RequestPreflight.Result result =
        RequestPreflight.verify(
            request,
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(request, temporaryDirectory, emptyRun)));

    assertFalse(result.preparedRequest().isPresent());
    assertEquals(
        List.of(
            "rich-text run file does not exist: missing-run.txt",
            "rich-text run must not be empty"),
        result.problems().stream().map(GridGrindProblemDetail.Problem::message).toList());
  }

  @Test
  void recordsUnexpectedPerStepResolutionFailuresAndKeepsResultInvariantsHonest() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("set-value"),
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetCell(new CellValueSpec.NumberValue(1.0)))));
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            temporaryDirectory,
            temporaryDirectory.resolve("scratch"),
            ExecutionGrantTestSupport.forPlan(request, temporaryDirectory));

    RequestPreflight.Result failed =
        RequestPreflight.verify(
            request,
            bindings,
            (step, ignoredBindings) -> {
              throw new IllegalStateException("resolution infrastructure unavailable");
            });

    assertFalse(failed.preparedRequest().isPresent());
    assertInstanceOf(ProblemContext.ResolveInputs.class, failed.problems().getFirst().context());
    assertThrows(IllegalStateException.class, failed::requirePreparedRequest);
    assertEquals(request, new RequestPreflight.PreparedRequest(request, bindings).request());
    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestPreflight.Result(java.util.Optional.empty(), List.of(), List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RequestPreflight.Result(
                java.util.Optional.of(new RequestPreflight.PreparedRequest(request, bindings)),
                failed.problems(),
                List.of()));
  }

  @Test
  void reportsOperationAndHostAcceptanceAuthorityDenialsThroughOnePreflightResult() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("ensure-budget"),
                    new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
                    new CellMutationAction.SetCell(new CellValueSpec.NumberValue(1.0d)))));
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            temporaryDirectory,
            temporaryDirectory.resolve("scratch"),
            new dev.erst.gridgrind.core.grant.ExecutionGrant(
                List.of(),
                List.of(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.WorkbookTargetAuthority
                    .WorkbookWide(),
                new dev.erst.gridgrind.core.grant.ExecutionGrant.PublicationAuthority.None(),
                List.of(),
                new dev.erst.gridgrind.core.grant.AcceptancePolicy.RequireAll(
                    List.of(
                        new dev.erst.gridgrind.core.grant.AcceptancePolicy.Requirement
                            .PreserveOpaqueOoxmlParts(List.of("/customXml/item1.xml"))))));

    RequestPreflight.Result result = RequestPreflight.verify(request, bindings);

    assertFalse(result.preparedRequest().isPresent());
    assertEquals(
        List.of(GridGrindProblemCode.AUTHORITY_DENIED, GridGrindProblemCode.INVALID_REQUEST),
        result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
    assertTrue(
        result.problems().stream()
            .map(GridGrindProblemDetail.Problem::context)
            .allMatch(ProblemContext.ValidateRequest.class::isInstance));
  }

  @Test
  void warnsForContainedAbsolutePersistencePathsAndSuppressesInvalidOverwriteOutputBinding()
      throws Exception {
    Path absoluteOutput = temporaryDirectory.resolve("output.xlsx");
    WorkbookPlan saveAs =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                absoluteOutput.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    RequestPreflight.Result saveAsResult =
        RequestPreflight.verify(
            saveAs,
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(saveAs, temporaryDirectory)));
    try {
      assertTrue(saveAsResult.problems().isEmpty());
      assertEquals(
          dev.erst.gridgrind.core.problem.GridGrindWarningCode.NON_PORTABLE_ABSOLUTE_PATH,
          saveAsResult.warnings().getFirst().code());
    } finally {
      saveAsResult.release();
    }

    WorkbookPlan invalidOverwrite =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.Overwrite(
                dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec.none()),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    RequestPreflight.Result invalidOverwriteResult =
        RequestPreflight.verify(
            invalidOverwrite,
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(invalidOverwrite, temporaryDirectory)));
    try {
      assertTrue(invalidOverwriteResult.problems().isEmpty());
      assertTrue(invalidOverwriteResult.preparedRequest().isPresent());
    } finally {
      invalidOverwriteResult.release();
    }

    assertInstanceOf(RequestPreflightPaths.class, RequestPreflightPaths.newForVerification());
    WorkbookPlan existingOverwrite =
        plan(
            new WorkbookPlan.WorkbookSource.ExistingFile("source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.Overwrite(
                dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec.none()));
    WorkbookPlan newOverwrite =
        plan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.Overwrite(
                dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec.none()));

    RequestPreflightPaths.PersistenceTarget target =
        RequestPreflightPaths.persistenceTarget(existingOverwrite).orElseThrow();
    assertEquals("source.xlsx", target.path());
    assertEquals(
        dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
        target.disposition());
    assertTrue(RequestPreflightPaths.persistenceTarget(newOverwrite).isEmpty());
  }

  @Test
  void classifiesCallerCorrectableSourceAndOutputPathFailuresWithoutPrivateTempPaths()
      throws Exception {
    Files.createDirectory(temporaryDirectory.resolve("source-directory.xlsx"));
    Files.createDirectory(temporaryDirectory.resolve("output-directory.xlsx"));
    Files.createFile(temporaryDirectory.resolve("existing-output.xlsx"));
    Files.writeString(temporaryDirectory.resolve("corrupt-source.xlsx"), "not an OOXML package");
    WorkbookPlan sourceDirectoryPlan =
        plan(
            new WorkbookPlan.WorkbookSource.ExistingFile("source-directory.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None());
    WorkbookPlan outputDirectoryPlan =
        plan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "output-directory.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT));
    WorkbookPlan outputCollisionPlan =
        plan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "existing-output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT));
    WorkbookPlan corruptSourcePlan =
        plan(
            new WorkbookPlan.WorkbookSource.ExistingFile("corrupt-source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None());

    RequestPreflight.Result sourceDirectory =
        RequestPreflight.verify(
            sourceDirectoryPlan,
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(sourceDirectoryPlan, temporaryDirectory)));
    RequestPreflight.Result outputDirectory =
        RequestPreflight.verify(
            outputDirectoryPlan,
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(outputDirectoryPlan, temporaryDirectory)));
    RequestPreflight.Result outputCollision =
        RequestPreflight.verify(
            outputCollisionPlan,
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(outputCollisionPlan, temporaryDirectory)));
    RequestPreflight.Result corruptSource =
        RequestPreflight.verify(
            corruptSourcePlan,
            new ExecutionInputBindings(
                temporaryDirectory,
                temporaryDirectory.resolve("scratch"),
                ExecutionGrantTestSupport.forPlan(corruptSourcePlan, temporaryDirectory)));

    assertEquals(
        GridGrindProblemCode.SOURCE_PATH_IS_DIRECTORY,
        sourceDirectory.problems().getFirst().code());
    assertEquals(
        GridGrindProblemCode.OUTPUT_PATH_IS_DIRECTORY,
        outputDirectory.problems().getFirst().code());
    assertEquals(
        GridGrindProblemCode.OUTPUT_PATH_ALREADY_EXISTS,
        outputCollision.problems().getFirst().code());
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_OPENABLE, corruptSource.problems().getFirst().code());
    assertTrue(
        List.of(sourceDirectory, outputDirectory, outputCollision, corruptSource).stream()
            .flatMap(result -> result.problems().stream())
            .map(GridGrindProblemDetail.Problem::message)
            .noneMatch(message -> message.contains("gridgrind-source-workbook-")));
  }

  private static WorkbookPlan plan(
      WorkbookPlan.WorkbookSource source, WorkbookPlan.WorkbookPersistence persistence) {
    return WorkbookPlan.standard(
        source,
        persistence,
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        List.of());
  }

  @Test
  void neverMasksAnEarlierPreflightProblemWithCleanupFailure() {
    AtomicBoolean attemptedClose = new AtomicBoolean();

    RequestPreflight.closeQuietly(
        () -> {
          attemptedClose.set(true);
          throw new java.io.IOException("private cleanup failed");
        });

    assertTrue(attemptedClose.get());
  }
}
