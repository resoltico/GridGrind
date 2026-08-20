package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies static request validation retains every independently provable authored violation. */
class StaticRequestValidatorTest {
  private final StaticRequestValidator validator = new StaticRequestValidator();

  @Test
  void retainsEquivalentStreamingWriteFailuresInAuthoredStepOrder() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.mode(new ExecutionModeInput.StreamingWrite()),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "ensure-sheet",
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.EnsureSheet()),
                setCell("set-first", "A1"),
                setCell("set-second", "A2")));

    List<String> messages =
        validator.validate(request).stream().map(GridGrindProblemDetail.Problem::message).toList();
    String unsupportedSetCell =
        GridGrindExecutionModeMetadata.streamingWrite().unsupportedActionMessage("SET_CELL");

    assertEquals(List.of(unsupportedSetCell, unsupportedSetCell), messages);
  }

  @Test
  void reportsEveryIndependentlyBoundOperationTargetMismatch() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "set-cell",
                    new WorkbookSelector.Current(),
                    new CellMutationAction.SetCell(new CellInput.NumberValue(1.0))),
                new MutationStep(
                    "ensure-sheet",
                    new WorkbookSelector.Current(),
                    new WorkbookMutationAction.EnsureSheet())));

    List<GridGrindProblemDetail.Problem> problems = validator.validate(request);

    assertEquals(2, problems.size());
    assertEquals(
        List.of("SET_CELL", "ENSURE_SHEET"),
        problems.stream()
            .map(GridGrindProblemDetail.Problem::message)
            .map(message -> message.substring(0, message.indexOf(" requires target type")))
            .toList());
    assertEquals(
        List.of(GridGrindProblemCode.INVALID_REQUEST, GridGrindProblemCode.INVALID_REQUEST),
        problems.stream().map(GridGrindProblemDetail.Problem::code).toList());
  }

  @Test
  void retainsBoundSiblingStaticFindingsButSuppressesRulesBelowMalformedFragments() {
    var analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [
                {
                  "stepId": "set-cell",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "action": {
                    "type": "SET_CELL",
                    "value": { "type": "NUMBER", "number": 1.0 }
                  }
                },
                {
                  "stepId": "missing-target",
                  "action": { "type": "ENSURE_SHEET" }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    List<GridGrindProblemDetail.Problem> problems =
        validator.validate(analysis, ProblemContextRequestSurfaces.RequestInput.standardInput());

    assertEquals(2, problems.size());
    assertEquals(
        List.of(
            "Missing required field 'steps[1].target'",
            "SET_CELL requires target type CELL_BY_ADDRESS or TABLE_CELL_BY_COLUMN_NAME but got WORKBOOK_CURRENT"),
        problems.stream().map(GridGrindProblemDetail.Problem::message).toList());
    assertInstanceOf(ProblemContext.ReadRequest.class, problems.getFirst().context());
    ProblemContext.ValidateRequest targetViolation =
        assertInstanceOf(ProblemContext.ValidateRequest.class, problems.get(1).context());
    assertEquals(
        "steps[0].target.type", targetViolation.json().orElseThrow().jsonPathValue().orElseThrow());
  }

  @Test
  void validatesBoundModeIncompatibilitiesDespiteAnUnrelatedMalformedStep() {
    var analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "EXISTING", "path": "source.xlsx" },
              "persistence": { "type": "NONE" },
              "execution": { "mode": { "type": "EVENT_READ" } },
              "steps": [
                {
                  "stepId": "malformed",
                  "action": { "type": "ENSURE_SHEET" }
                },
                {
                  "stepId": "mutate",
                  "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                  "action": {
                    "type": "SET_CELL",
                    "value": { "type": "NUMBER", "number": 1.0 }
                  }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    List<GridGrindProblemDetail.Problem> problems =
        validator.validate(analysis, ProblemContextRequestSurfaces.RequestInput.standardInput());

    assertEquals(2, problems.size());
    assertEquals("Missing required field 'steps[0].target'", problems.getFirst().message());
    ProblemContext.ValidateRequest modeViolation =
        assertInstanceOf(ProblemContext.ValidateRequest.class, problems.get(1).context());
    assertEquals(
        "execution.mode", modeViolation.json().orElseThrow().jsonPathValue().orElseThrow());
    assertTrue(problems.get(1).message().contains("supports inspection steps only"));
  }

  @Test
  void validatesPersistenceAsSoonAsItsOwnRootDependenciesBind() {
    var analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "OVERWRITE" },
              "steps": [
                {
                  "stepId": "missing-target",
                  "action": { "type": "ENSURE_SHEET" }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    List<GridGrindProblemDetail.Problem> problems =
        validator.validate(analysis, ProblemContextRequestSurfaces.RequestInput.standardInput());

    assertEquals(
        List.of(
            "Missing required field 'steps[0].target'",
            "OVERWRITE persistence requires an EXISTING source; a NEW workbook has no source file to overwrite"),
        problems.stream().map(GridGrindProblemDetail.Problem::message).toList());
    ProblemContext.ValidateRequest persistenceViolation =
        assertInstanceOf(ProblemContext.ValidateRequest.class, problems.get(1).context());
    assertEquals(
        "persistence.type",
        persistenceViolation.json().orElseThrow().jsonPathValue().orElseThrow());
  }

  @Test
  void validatesCalculationOrderingWithoutAnUnrelatedSourceFragment() {
    var analysis =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": null,
              "persistence": { "type": "NONE" },
              "execution": {
                "calculation": { "strategy": { "type": "EVALUATE_ALL" } }
              },
              "steps": [
                {
                  "stepId": "inspect-workbook",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "query": { "type": "GET_WORKBOOK_SUMMARY" }
                },
                {
                  "stepId": "ensure-sheet",
                  "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                  "action": { "type": "ENSURE_SHEET" }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));

    List<GridGrindProblemDetail.Problem> problems =
        validator.validate(analysis, ProblemContextRequestSurfaces.RequestInput.standardInput());

    assertEquals(2, problems.size());
    assertEquals(
        "Field 'source' must be omitted when absent; explicit null is not accepted.",
        problems.getFirst().message());
    assertEquals(
        "execution.calculation.strategy=EVALUATE_ALL requires all MUTATION steps to appear before any ASSERTION or INSPECTION step so calculation can run once at the mutation-to-observation boundary",
        problems.get(1).message());
    ProblemContext.ValidateRequest calculationViolation =
        assertInstanceOf(ProblemContext.ValidateRequest.class, problems.get(1).context());
    assertEquals(
        "execution.calculation",
        calculationViolation.json().orElseThrow().jsonPathValue().orElseThrow());
  }

  @Test
  void toleratesPartialFragmentsAndPreservesKnownSaveAsRequestShape() {
    var missingRootFragments =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "steps": [
                {
                  "stepId": "set-cell",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "action": { "type": "SET_CELL", "value": { "type": "NUMBER", "number": 1.0 } }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    var missingExecution =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "execution": null,
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    var missingPersistence =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "steps": [
                {
                  "stepId": "set-cell",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "action": { "type": "SET_CELL", "value": { "type": "NUMBER", "number": 1.0 } }
                }
              ]
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    var missingSteps =
        GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": null
            }
            """
                .getBytes(StandardCharsets.UTF_8));
    WorkbookPlan saveAs =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "set-cell",
                    new WorkbookSelector.Current(),
                    new CellMutationAction.SetCell(new CellInput.NumberValue(1.0)))));

    List<GridGrindProblemDetail.Problem> rootProblems =
        validator.validate(
            missingRootFragments, ProblemContextRequestSurfaces.RequestInput.standardInput());

    assertEquals(3, rootProblems.size());
    ProblemContext.ValidateRequest targetViolation =
        assertInstanceOf(ProblemContext.ValidateRequest.class, rootProblems.getLast().context());
    assertInstanceOf(
        ProblemContextRequestSurfaces.RequestShape.Unknown.class, targetViolation.request());
    assertEquals(
        1,
        validator
            .validate(missingExecution, ProblemContextRequestSurfaces.RequestInput.standardInput())
            .size());
    assertEquals(
        1,
        validator
            .validate(missingSteps, ProblemContextRequestSurfaces.RequestInput.standardInput())
            .size());
    assertEquals(
        2,
        validator
            .validate(
                missingPersistence, ProblemContextRequestSurfaces.RequestInput.standardInput())
            .size());
    ProblemContext.ValidateRequest saveAsViolation =
        assertInstanceOf(
            ProblemContext.ValidateRequest.class, validator.validate(saveAs).getFirst().context());
    assertEquals("SAVE_AS", saveAsViolation.request().persistenceTypeValue().orElseThrow());
    assertTrue(saveAsViolation.json().isEmpty());
  }

  private static MutationStep setCell(String stepId, String address) {
    return new MutationStep(
        stepId,
        new CellSelector.ByAddress("Budget", address),
        new CellMutationAction.SetCell(new CellInput.NumberValue(1.0)));
  }
}
