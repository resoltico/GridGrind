package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.ResponseOutput;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.PersistenceReference;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.WorkbookReference;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests the total, serialized-metadata-free ordering of protocol diagnostics. */
class DiagnosticOrderTest {
  private static final RequestShape SHAPE = RequestShape.known("NEW", "NONE");

  @Test
  void ordersEveryProblemContextByPipelinePhaseAndStablePhaseOrdinal() {
    GridGrindProblemDetail.Problem writeResponse =
        problem(new ProblemContext.WriteResponse(ResponseOutput.standardOutput()));
    GridGrindProblemDetail.Problem calculationPreflight =
        problem(new ProblemContext.ExecuteCalculation.Preflight(SHAPE, ProblemLocation.unknown()));
    GridGrindProblemDetail.Problem validation = problem(new ProblemContext.ValidateRequest(SHAPE));
    GridGrindProblemDetail.Problem parseArguments =
        problem(new ProblemContext.ParseArguments(CliArgument.named("--request")));
    GridGrindProblemDetail.Problem cliRuntime = problem(new CliRuntimeContext());
    GridGrindProblemDetail.Problem readRequest =
        problem(
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.unavailable()));
    GridGrindProblemDetail.Problem bindRequest =
        problem(
            new ProblemContext.BindRequest(
                RequestInput.standardInput(), JsonLocation.unavailable()));
    GridGrindProblemDetail.Problem resolveInputs =
        problem(new ProblemContext.ResolveInputs(SHAPE, InputReference.unknown()));
    GridGrindProblemDetail.Problem openWorkbook =
        problem(new ProblemContext.OpenWorkbook(SHAPE, WorkbookReference.newWorkbook()));
    GridGrindProblemDetail.Problem calculationExecution =
        problem(new ProblemContext.ExecuteCalculation.Execution(SHAPE, ProblemLocation.unknown()));
    GridGrindProblemDetail.Problem executeStep =
        problem(
            new ProblemContext.ExecuteStep(
                SHAPE,
                new StepReference(2, "step-2", "MUTATION", "SET_CELL"),
                ProblemLocation.unknown()));
    GridGrindProblemDetail.Problem persistWorkbook =
        problem(new ProblemContext.PersistWorkbook(SHAPE, PersistenceReference.saveAs("out.xlsx")));
    GridGrindProblemDetail.Problem executeRequest =
        problem(new ProblemContext.ExecuteRequest(SHAPE));

    assertEquals(
        List.of(
            parseArguments,
            cliRuntime,
            readRequest,
            bindRequest,
            validation,
            calculationPreflight,
            resolveInputs,
            openWorkbook,
            writeResponse,
            calculationExecution,
            persistWorkbook,
            executeRequest,
            executeStep),
        DiagnosticOrder.problems(
            List.of(
                writeResponse,
                calculationPreflight,
                validation,
                parseArguments,
                cliRuntime,
                readRequest,
                bindRequest,
                resolveInputs,
                openWorkbook,
                calculationExecution,
                executeStep,
                persistWorkbook,
                executeRequest)));
  }

  @Test
  void ordersRequestLocationsByOffsetThenStepThenDuplicateOccurrence() {
    GridGrindProblemDetail.Problem unpositioned = readProblem(JsonLocation.pathOnly("source"));
    GridGrindProblemDetail.Problem malformedStepIndex =
        readProblem(JsonLocation.pathOnly("steps[nope].stepId"));
    GridGrindProblemDetail.Problem unclosedStepIndex =
        readProblem(JsonLocation.pathOnly("steps[2.stepId"));
    GridGrindProblemDetail.Problem secondStep =
        readProblem(JsonLocation.pathOnly("steps[2].stepId"));
    GridGrindProblemDetail.Problem fourthStep =
        readProblem(JsonLocation.pathOnly("steps[4].stepId"));
    GridGrindProblemDetail.Problem lateOffset =
        readProblem(JsonLocation.pathAtByteOffset("steps[4].stepId", 10));
    GridGrindProblemDetail.Problem earlyOffset =
        readProblem(JsonLocation.pathAtByteOffset("steps[5].stepId", 5));
    GridGrindProblemDetail.Problem secondDuplicate =
        readProblem(JsonLocation.duplicateKey("steps[2]", "stepId", 2, 7));
    GridGrindProblemDetail.Problem firstDuplicate =
        readProblem(JsonLocation.duplicateKey("steps[2]", "stepId", 1, 7));

    assertEquals(
        List.of(
            earlyOffset,
            firstDuplicate,
            secondDuplicate,
            lateOffset,
            unpositioned,
            malformedStepIndex,
            unclosedStepIndex,
            secondStep,
            fourthStep),
        DiagnosticOrder.problems(
            List.of(
                unpositioned,
                malformedStepIndex,
                unclosedStepIndex,
                secondStep,
                fourthStep,
                lateOffset,
                earlyOffset,
                secondDuplicate,
                firstDuplicate)));
  }

  @Test
  void ordersStructuralIntakeBeforeAnEarlierLocatedBindingFailure() {
    GridGrindProblemDetail.Problem bindingFailure =
        problem(
            new ProblemContext.BindRequest(
                RequestInput.standardInput(), JsonLocation.pathAtByteOffset("source.path", 2)));
    GridGrindProblemDetail.Problem structuralFailure =
        problem(
            new ProblemContext.ReadRequest(
                RequestInput.standardInput(), JsonLocation.pathAtByteOffset("unexpected", 40)));

    assertEquals(
        List.of(structuralFailure, bindingFailure),
        DiagnosticOrder.problems(List.of(bindingFailure, structuralFailure)));
  }

  @Test
  void ordersWarningsByStepThenCodeWithoutExposingOrderingMetadata() {
    RequestWarning later = warning(3, "later");
    RequestWarning earlier = warning(1, "earlier");
    RequestWarning firstSameStep = warning(2, "first-same-step");
    RequestWarning secondSameStep = warning(2, "second-same-step");

    assertEquals(List.of(earlier, later), DiagnosticOrder.warnings(List.of(later, earlier)));
    assertEquals(
        List.of(firstSameStep, secondSameStep),
        DiagnosticOrder.warnings(List.of(firstSameStep, secondSameStep)));
  }

  @Test
  void ordersRequestPathWarningsBeforeStepWarningsAndModelsBothLocationVariants() {
    RequestWarning requestPath =
        RequestWarning.nonPortableAbsolutePath("/work/report.xlsx", "persistence");
    RequestWarning step = warning(0, "step");

    assertEquals(List.of(requestPath, step), DiagnosticOrder.warnings(List.of(step, requestPath)));
    RequestWarningLocation requestPathLocation =
        new RequestWarningLocation.RequestPath("/work/report.xlsx", "persistence");
    RequestWarningLocation stepLocation = new RequestWarningLocation.Step(0, "step", "SET_CELL");
    assertEquals(-1, requestPathLocation.orderingStepIndex());
    assertEquals(0, stepLocation.orderingStepIndex());
    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestWarningLocation.Step(-1, "step", "SET_CELL"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestWarningLocation.RequestPath(" ", "persistence"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RequestWarningLocation.RequestPath("/work/report.xlsx", " "));
  }

  private static GridGrindProblemDetail.Problem problem(ProblemContext context) {
    return GridGrindProblemDetail.Problem.of(
        GridGrindProblemCode.INVALID_REQUEST, "problem", context);
  }

  private static GridGrindProblemDetail.Problem readProblem(JsonLocation location) {
    return problem(new ProblemContext.ReadRequest(RequestInput.standardInput(), location));
  }

  private static RequestWarning warning(int stepIndex, String stepId) {
    return new RequestWarning(
        GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
        stepIndex,
        stepId,
        "SET_CELL",
        "Quote the sheet name.");
  }
}
