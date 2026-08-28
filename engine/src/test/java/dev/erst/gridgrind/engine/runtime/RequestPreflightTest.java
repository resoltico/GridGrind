package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CellGridInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies phase-four source checks collect independent facts without workbook mutation. */
class RequestPreflightTest {
  @TempDir Path temporaryDirectory;

  @Test
  void collectsIndependentAssetAndWorkbookSourceFailuresWithoutAResolvedPlan() throws Exception {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("missing-source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "set-title",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(
                        new CellInput.Text(new TextSourceInput.Utf8File("missing-title.txt")))),
                new MutationStep(
                    "set-subtitle",
                    new CellSelector.ByAddress("Budget", "A2"),
                    new CellMutationAction.SetCell(
                        new CellInput.Text(
                            new TextSourceInput.Utf8File("missing-subtitle.txt"))))));
    Path tempRoot = temporaryDirectory.resolve("scratch");

    RequestPreflight.Result result =
        RequestPreflight.verify(request, new ExecutionInputBindings(temporaryDirectory, tempRoot));

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
                    problem.context().stage().endsWith("INPUTS")
                        || "OPEN_WORKBOOK".equals(problem.context().stage())));
  }

  @Test
  void doctorBatchesIndependentPreflightFailuresAlongsideStaticFindings() {
    var analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
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
                      "source": { "type": "UTF8_FILE", "path": "missing-title.txt" }
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
                    temporaryDirectory, temporaryDirectory.resolve("scratch")));

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
              "protocolVersion": "V2",
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
                        { "type": "TEXT", "source": { "type": "UTF8_FILE", "path": "missing-first.txt" } },
                        { "type": "TEXT", "source": { "type": "UTF8_FILE", "path": "missing-second.txt" } }
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
            new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch")),
            analysis);

    assertEquals(2, result.problems().size());
    assertEquals(
        List.of(
            "steps[0].action.rows.cells[0][0].source.path",
            "steps[0].action.rows.cells[0][1].source.path"),
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
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "set-values",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetRange(
                        new CellGridInput.Typed(
                            List.of(
                                List.of(
                                    new CellInput.Text(
                                        new TextSourceInput.Utf8File("missing-first.txt")),
                                    new CellInput.Text(
                                        new TextSourceInput.Utf8File("missing-second.txt")))))))));

    RequestPreflight.Result result =
        RequestPreflight.verify(
            request,
            new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch")));

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
    Files.createFile(temporaryDirectory.resolve("empty-run.txt"));
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "set-rich-text",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetRange(
                        new CellGridInput.Typed(
                            List.of(
                                List.of(
                                    new CellInput.RichText(
                                        List.of(
                                            new dev.erst.gridgrind.contract.dto.RichTextRunInput(
                                                new TextSourceInput.Utf8File("missing-run.txt"),
                                                java.util.Optional.empty()))),
                                    new CellInput.RichText(
                                        List.of(
                                            new dev.erst.gridgrind.contract.dto.RichTextRunInput(
                                                new TextSourceInput.Utf8File("empty-run.txt"),
                                                java.util.Optional.empty()))))))))));

    RequestPreflight.Result result =
        RequestPreflight.verify(
            request,
            new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch")));

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
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "set-value",
                    new CellSelector.ByAddress("Budget", "A1"),
                    new CellMutationAction.SetCell(new CellInput.NumberValue(1.0)))));
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch"));

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
  void warnsForContainedAbsolutePersistencePathsAndSuppressesInvalidOverwriteOutputBinding()
      throws Exception {
    Path absoluteOutput = temporaryDirectory.resolve("output.xlsx");
    WorkbookPlan saveAs =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                absoluteOutput.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    RequestPreflight.Result saveAsResult =
        RequestPreflight.verify(
            saveAs,
            new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch")));
    try {
      assertTrue(saveAsResult.problems().isEmpty());
      assertEquals(
          dev.erst.gridgrind.contract.dto.GridGrindWarningCode.NON_PORTABLE_ABSOLUTE_PATH,
          saveAsResult.warnings().getFirst().code());
    } finally {
      saveAsResult.release();
    }

    WorkbookPlan invalidOverwrite =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.Overwrite(
                dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput.none()),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    RequestPreflight.Result invalidOverwriteResult =
        RequestPreflight.verify(
            invalidOverwrite,
            new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch")));
    try {
      assertTrue(invalidOverwriteResult.problems().isEmpty());
      assertTrue(invalidOverwriteResult.preparedRequest().isPresent());
    } finally {
      invalidOverwriteResult.release();
    }
  }

  @Test
  void batchesFormulaInputsAndPreparesOverwriteTargetsFromExistingSources() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("missing-source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.Overwrite(
                dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput.none()),
            ExecutionPolicyInput.defaults(),
            new FormulaEnvironmentInput(
                List.of(new FormulaExternalWorkbookInput("External.xlsx", "missing-external.xlsx")),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()),
            List.of());

    RequestPreflight.Result result =
        RequestPreflight.verify(
            request,
            new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch")));

    assertFalse(result.preparedRequest().isPresent());
    assertEquals(
        List.of(
            GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND, GridGrindProblemCode.WORKBOOK_NOT_FOUND),
        result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
  }

  @Test
  void rejectsColumnEditsAgainstExistingSourcesThatAlreadyContainFormulas() throws Exception {
    Path sourcePath = temporaryDirectory.resolve("existing-formulas.xlsx");
    try (XSSFWorkbook source = new XSSFWorkbook()) {
      source.createSheet("Ops").createRow(0).createCell(0).setCellFormula("1+1");
      try (var output = Files.newOutputStream(sourcePath)) {
        source.write(output);
      }
    }
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.getFileName().toString()),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "insert-column",
                    new ColumnBandSelector.Insertion("Ops", 1, 1),
                    new WorkbookMutationAction.InsertColumns())));

    RequestPreflight.Result result =
        RequestPreflight.verify(
            request,
            new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch")));

    assertFalse(result.preparedRequest().isPresent());
    GridGrindProblemDetail.Problem problem = result.problems().getFirst();
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, problem.code());
    assertEquals(
        "Column insert, delete, and shift operations must appear before formula authoring because"
            + " formula-bearing workbook column edits are unsupported.",
        problem.message());
    assertInstanceOf(ProblemContext.OpenWorkbook.class, problem.context());
    assertFalse(problem.message().contains("gridgrind-source-workbook-"));
  }

  @Test
  void classifiesCallerCorrectableSourceAndOutputPathFailuresWithoutPrivateTempPaths()
      throws Exception {
    Files.createDirectory(temporaryDirectory.resolve("source-directory.xlsx"));
    Files.createDirectory(temporaryDirectory.resolve("output-directory.xlsx"));
    Files.createFile(temporaryDirectory.resolve("existing-output.xlsx"));
    Files.writeString(temporaryDirectory.resolve("corrupt-source.xlsx"), "not an OOXML package");
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch"));

    RequestPreflight.Result sourceDirectory =
        RequestPreflight.verify(
            plan(
                new WorkbookPlan.WorkbookSource.ExistingFile("source-directory.xlsx"),
                new WorkbookPlan.WorkbookPersistence.None()),
            bindings);
    RequestPreflight.Result outputDirectory =
        RequestPreflight.verify(
            plan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    "output-directory.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT)),
            bindings);
    RequestPreflight.Result outputCollision =
        RequestPreflight.verify(
            plan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    "existing-output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT)),
            bindings);
    RequestPreflight.Result corruptSource =
        RequestPreflight.verify(
            plan(
                new WorkbookPlan.WorkbookSource.ExistingFile("corrupt-source.xlsx"),
                new WorkbookPlan.WorkbookPersistence.None()),
            bindings);

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
        ExecutionPolicyInput.defaults(),
        FormulaEnvironmentInput.empty(),
        List.of());
  }

  @Test
  void preservesFormulaWorkbookPathEscapeDiagnostics() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            new FormulaEnvironmentInput(
                List.of(new FormulaExternalWorkbookInput("External.xlsx", "../outside.xlsx")),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()),
            List.of());

    RequestPreflight.Result result =
        RequestPreflight.verify(
            request,
            new ExecutionInputBindings(temporaryDirectory, temporaryDirectory.resolve("scratch")));

    assertEquals(
        List.of(GridGrindProblemCode.PATH_ESCAPES_ROOT),
        result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
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
