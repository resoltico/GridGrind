package dev.erst.gridgrind.engine.api;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;

/** Coverage for the narrow engine API seam exported to downstream transports. */
class GridGrindEngineApiTest {
  @Test
  void requestInputsNormalizeAndDefensivelyCopyBoundStandardInput() {
    byte[] original = new byte[] {1, 2, 3};
    GridGrindRequestInputs inputs =
        new GridGrindRequestInputs(Path.of("tmp", "..", "tmp", "inputs"), original);

    original[0] = 9;
    byte[] firstRead = inputs.standardInputBytes().orElseThrow();
    firstRead[1] = 8;

    assertTrue(inputs.workingDirectory().isAbsolute());
    assertTrue(inputs.workingDirectory().endsWith(Path.of("tmp", "inputs")));
    assertTrue(inputs.hasStandardInput());
    assertArrayEquals(new byte[] {1, 2, 3}, inputs.standardInputBytes().orElseThrow());
  }

  @Test
  void requestInputsWithoutStandardInputUseNormalizedWorkingDirectories() {
    GridGrindRequestInputs explicit = new GridGrindRequestInputs(Path.of("."));
    GridGrindRequestInputs processDefault = GridGrindRequestInputs.processDefault();

    assertTrue(explicit.workingDirectory().isAbsolute());
    assertFalse(explicit.hasStandardInput());
    assertTrue(processDefault.workingDirectory().isAbsolute());
    assertFalse(processDefault.hasStandardInput());
  }

  @Test
  void requestExecutorDefaultOverloadsUseExpectedDefaultsAndRejectNullDelegates() {
    WorkbookPlan request = GridGrindProtocolCatalog.requestTemplate();
    AtomicReference<GridGrindRequestInputs> observedInputs = new AtomicReference<>();
    AtomicReference<GridGrindJournalSink> observedSink = new AtomicReference<>();
    GridGrindResponse success = GridGrindResponses.success(List.of(), List.of(), List.of());
    GridGrindRequestExecutor executor =
        (observedRequest, inputs, sink) -> {
          assertSame(request, observedRequest);
          observedInputs.set(inputs);
          observedSink.set(sink);
          return success;
        };

    assertSame(executor, GridGrindRequestExecutor.requireNonNull(executor));
    assertThrows(NullPointerException.class, () -> GridGrindRequestExecutor.requireNonNull(null));

    GridGrindRequestInputs boundInputs =
        new GridGrindRequestInputs(Path.of("engine-api-inputs"), new byte[] {4, 5});
    assertSame(success, executor.execute(request, boundInputs));
    assertSame(boundInputs, observedInputs.get());
    assertSame(GridGrindJournalSink.NOOP, observedSink.get());

    GridGrindJournalSink sink = event -> {};
    assertSame(success, executor.execute(request, sink));
    assertTrue(observedInputs.get().workingDirectory().isAbsolute());
    assertFalse(observedInputs.get().hasStandardInput());
    assertSame(sink, observedSink.get());

    assertSame(success, executor.execute(request));
    assertTrue(observedInputs.get().workingDirectory().isAbsolute());
    assertFalse(observedInputs.get().hasStandardInput());
    assertSame(GridGrindJournalSink.NOOP, observedSink.get());
  }

  @Test
  void journalSinkAndDoctorNullGuardsStayCovered() {
    GridGrindJournalSink sink = event -> {};
    GridGrindJournalSink.NOOP.emit(null);
    assertSame(sink, GridGrindJournalSink.requireNonNull(sink));
    assertThrows(NullPointerException.class, () -> GridGrindJournalSink.requireNonNull(null));

    GridGrindRequestDoctor doctor =
        new GridGrindRequestDoctor() {
          @Override
          public RequestDoctorReport diagnose(WorkbookPlan request) {
            return cleanDoctorSummary();
          }

          @Override
          public RequestDoctorReport diagnose(WorkbookPlan request, GridGrindRequestInputs inputs) {
            return cleanDoctorSummary();
          }
        };
    assertSame(doctor, GridGrindRequestDoctor.requireNonNull(doctor));
    assertThrows(NullPointerException.class, () -> GridGrindRequestDoctor.requireNonNull(null));
  }

  @Test
  void productionEngineAdaptersRequestRequirementsAndProblemHelpersStayCovered()
      throws IOException {
    WorkbookPlan template = GridGrindProtocolCatalog.requestTemplate();
    WorkbookPlan standardInputRequest = standardInputRequest();

    assertFalse(GridGrindRequestRequirements.requiresStandardInput(template));
    assertTrue(GridGrindRequestRequirements.requiresStandardInput(standardInputRequest));

    GridGrindRequestExecutor executor = GridGrindEngine.requestExecutor();
    GridGrindRequestDoctor doctor = GridGrindEngine.requestDoctor();
    assertNotNull(executor);
    assertNotNull(doctor);

    assertInstanceOf(GridGrindResponse.Success.class, executor.execute(template));
    assertInstanceOf(
        GridGrindResponse.Success.class,
        executor.execute(
            template,
            new GridGrindRequestInputs(Path.of("engine-api-runtime"), new byte[] {7}),
            GridGrindJournalSink.NOOP));

    RequestDoctorReport clean = doctor.diagnose(template);
    RequestDoctorReport cleanWithInputs =
        doctor.diagnose(template, new GridGrindRequestInputs(Path.of("engine-api-doctor")));
    assertTrue(clean.valid());
    assertTrue(cleanWithInputs.valid());

    ProblemContext context =
        new ProblemContext.ValidateRequest(
            ProblemContextRequestSurfaces.RequestShape.known("NEW", "NONE"));
    GridGrindProblemDetail.Problem fromException =
        GridGrindProblems.fromException(new IllegalArgumentException("boom"), context);
    GridGrindProblemDetail.Problem explicitProblem =
        GridGrindProblems.problem(
            GridGrindProblemCode.INVALID_REQUEST,
            "broken request",
            context,
            new IllegalStateException("broken"));
    GridGrindProblemDetail.ProblemCause causeFromProblem =
        GridGrindProblems.problemCause(fromException);
    GridGrindProblemDetail.Problem withStructuredCause =
        GridGrindProblems.problem(
            GridGrindProblemCode.INVALID_REQUEST,
            "broken request",
            context,
            List.of(causeFromProblem));
    GridGrindProblemDetail.Problem appended =
        GridGrindProblems.appendCause(explicitProblem, causeFromProblem);

    assertSame(context, fromException.context());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, explicitProblem.code());
    assertEquals("VALIDATE_REQUEST", causeFromProblem.stage());
    assertEquals(1, withStructuredCause.causes().size());
    assertEquals(2, appended.causes().size());
  }

  @Test
  void productionRequestExecutorForwardsVerboseJournalEventsToThePublicSink() throws IOException {
    WorkbookPlan verboseRequest = verboseRequest();
    AtomicReference<dev.erst.gridgrind.contract.dto.ExecutionJournal.Event> observedEvent =
        new AtomicReference<>();

    GridGrindResponse response =
        GridGrindEngine.requestExecutor()
            .execute(
                verboseRequest,
                new GridGrindRequestInputs(Path.of("engine-api-verbose")),
                observedEvent::set);

    assertInstanceOf(GridGrindResponse.Success.class, response);
    assertNotNull(observedEvent.get());
    assertEquals("PLAN", observedEvent.get().category());
  }

  private static RequestDoctorReport cleanDoctorSummary() {
    return RequestDoctorReport.clean(
        new RequestDoctorReport.Summary(
            "NEW", "NONE", "FULL_XSSF", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0));
  }

  private static WorkbookPlan standardInputRequest() throws IOException {
    return GridGrindJson.readRequest(
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "execution": {
            "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
            "journal": { "level": "NORMAL" },
            "calculation": {
              "strategy": { "type": "DO_NOT_CALCULATE" },
              "markRecalculateOnOpen": false
            }
          },
          "formulaEnvironment": {
            "externalWorkbooks": [],
            "missingWorkbookPolicy": "ERROR",
            "udfToolpacks": []
          },
          "steps": [
            {
              "stepId": "set",
              "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
              "action": {
                "type": "SET_CELL",
                "value": { "type": "TEXT", "source": { "type": "STANDARD_INPUT" } }
              }
            }
          ]
        }
        """
            .getBytes(StandardCharsets.UTF_8));
  }

  private static WorkbookPlan verboseRequest() throws IOException {
    return GridGrindJson.readRequest(
        """
        {
          "protocolVersion": "V1",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "execution": {
            "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
            "journal": { "level": "VERBOSE" },
            "calculation": {
              "strategy": { "type": "DO_NOT_CALCULATE" },
              "markRecalculateOnOpen": false
            }
          },
          "formulaEnvironment": {
            "externalWorkbooks": [],
            "missingWorkbookPolicy": "ERROR",
            "udfToolpacks": []
          },
          "steps": []
        }
        """
            .getBytes(StandardCharsets.UTF_8));
  }
}
