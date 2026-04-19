package dev.erst.gridgrind.jazzer.tool;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.executor.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.jazzer.support.GeneratedProtocolWorkflow;
import dev.erst.gridgrind.jazzer.support.GridGrindFuzzData;
import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.OperationSequenceModel;
import dev.erst.gridgrind.jazzer.support.SequenceIntrospection;
import dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks;
import dev.erst.gridgrind.jazzer.support.XlsxRoundTripVerifier;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** Replays raw Jazzer inputs outside active fuzzing and classifies their semantic outcome. */
public final class JazzerReplaySupport {
  private JazzerReplaySupport() {}

  /** Returns the stable replay expectation captured from one replay outcome. */
  public static ReplayExpectation expectationFor(ReplayOutcome outcome) {
    Objects.requireNonNull(outcome, "outcome must not be null");
    return new ReplayExpectation(outcomeKind(outcome), outcome.details());
  }

  /** Replays one raw input against the selected harness and returns a structured outcome. */
  public static ReplayOutcome replay(JazzerHarness harness, byte[] input) {
    Objects.requireNonNull(harness, "harness must not be null");
    Objects.requireNonNull(input, "input must not be null");
    if (harness.equals(JazzerHarness.protocolRequest())) {
      return replayProtocolRequest(input);
    }
    if (harness.equals(JazzerHarness.protocolWorkflow())) {
      return replayProtocolWorkflow(input);
    }
    if (harness.equals(JazzerHarness.engineCommandSequence())) {
      return replayCommandSequence(input);
    }
    if (harness.equals(JazzerHarness.xlsxRoundTrip())) {
      return replayRoundTrip(input);
    }
    throw new IllegalArgumentException("Unsupported Jazzer harness: " + harness.key());
  }

  private static ReplayOutcome replayProtocolRequest(byte[] input) {
    try {
      WorkbookPlan request = GridGrindJson.readRequest(input);
      ProtocolRequestDetails details =
          new ProtocolRequestDetails(
              input.length,
              "PARSED",
              SequenceIntrospection.sourceKind(request),
              SequenceIntrospection.persistenceKind(request),
              request.mutationSteps().size(),
              SequenceIntrospection.mutationKinds(request.mutationSteps()),
              SequenceIntrospection.styleKinds(request.mutationSteps()),
              SequenceIntrospection.assertionCount(request),
              SequenceIntrospection.assertionKinds(request.assertionSteps()),
              SequenceIntrospection.inspectionCount(request),
              SequenceIntrospection.inspectionKinds(request.inspectionSteps()));
      return new ReplayOutcome.Success(JazzerHarness.protocolRequest().key(), details);
    } catch (InvalidJsonException expected) {
      ProtocolRequestDetails details =
          new ProtocolRequestDetails(
              input.length,
              "INVALID_JSON",
              "NOT_PARSED",
              "NOT_PARSED",
              0,
              Map.of(),
              Map.of(),
              0,
              Map.of(),
              0,
              Map.of());
      return new ReplayOutcome.ExpectedInvalid(
          JazzerHarness.protocolRequest().key(),
          expected.getClass().getSimpleName(),
          expected.getMessage(),
          details);
    } catch (InvalidRequestShapeException expected) {
      ProtocolRequestDetails details =
          new ProtocolRequestDetails(
              input.length,
              "INVALID_REQUEST_SHAPE",
              "NOT_PARSED",
              "NOT_PARSED",
              0,
              Map.of(),
              Map.of(),
              0,
              Map.of(),
              0,
              Map.of());
      return new ReplayOutcome.ExpectedInvalid(
          JazzerHarness.protocolRequest().key(),
          expected.getClass().getSimpleName(),
          expected.getMessage(),
          details);
    } catch (InvalidRequestException expected) {
      ProtocolRequestDetails details =
          new ProtocolRequestDetails(
              input.length,
              "INVALID_REQUEST",
              "NOT_PARSED",
              "NOT_PARSED",
              0,
              Map.of(),
              Map.of(),
              0,
              Map.of(),
              0,
              Map.of());
      return new ReplayOutcome.ExpectedInvalid(
          JazzerHarness.protocolRequest().key(),
          expected.getClass().getSimpleName(),
          expected.getMessage(),
          details);
    } catch (IOException unexpected) {
      ProtocolRequestDetails details =
          new ProtocolRequestDetails(
              input.length,
              "IO_ERROR",
              "NOT_PARSED",
              "NOT_PARSED",
              0,
              Map.of(),
              Map.of(),
              0,
              Map.of(),
              0,
              Map.of());
      return unexpectedFailure(JazzerHarness.protocolRequest(), unexpected, details);
    } catch (RuntimeException unexpected) {
      ProtocolRequestDetails details =
          new ProtocolRequestDetails(
              input.length,
              "RUNTIME_FAILURE",
              "NOT_PARSED",
              "NOT_PARSED",
              0,
              Map.of(),
              Map.of(),
              0,
              Map.of(),
              0,
              Map.of());
      return unexpectedFailure(JazzerHarness.protocolRequest(), unexpected, details);
    }
  }

  private static ReplayOutcome replayProtocolWorkflow(byte[] input) {
    GridGrindFuzzData data = GridGrindFuzzData.replay(input);
    GeneratedProtocolWorkflow workflow = null;
    WorkbookPlan request = null;
    GridGrindResponse response = null;
    try {
      workflow = OperationSequenceModel.nextProtocolWorkflow(data);
      request = workflow.request();
      response = new DefaultGridGrindRequestExecutor().execute(request);
      WorkbookInvariantChecks.requireResponseShape(response);
      WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response);
      ProtocolWorkflowDetails details = workflowDetails(input.length, request, response);
      return new ReplayOutcome.Success(JazzerHarness.protocolWorkflow().key(), details);
    } catch (IllegalArgumentException expected) {
      ProtocolWorkflowDetails details = workflowDetails(input.length, request, response);
      return new ReplayOutcome.ExpectedInvalid(
          JazzerHarness.protocolWorkflow().key(),
          expected.getClass().getSimpleName(),
          expected.getMessage(),
          details);
    } catch (IOException unexpected) {
      ProtocolWorkflowDetails details = workflowDetails(input.length, request, response);
      return unexpectedFailure(JazzerHarness.protocolWorkflow(), unexpected, details);
    } catch (RuntimeException unexpected) {
      ProtocolWorkflowDetails details = workflowDetails(input.length, request, response);
      return unexpectedFailure(JazzerHarness.protocolWorkflow(), unexpected, details);
    } finally {
      if (workflow != null) {
        workflow.cleanup();
      }
    }
  }

  private static ReplayOutcome replayCommandSequence(byte[] input) {
    GridGrindFuzzData data = GridGrindFuzzData.replay(input);
    List<WorkbookCommand> commands = List.of();
    try {
      commands = OperationSequenceModel.nextWorkbookCommands(data);
      try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
        new WorkbookCommandExecutor().apply(workbook, commands);
        WorkbookInvariantChecks.requireWorkbookShape(workbook);
      }
      CommandSequenceDetails details = commandSequenceDetails(input.length, commands);
      return new ReplayOutcome.Success(JazzerHarness.engineCommandSequence().key(), details);
    } catch (IllegalArgumentException expected) {
      CommandSequenceDetails details = commandSequenceDetails(input.length, commands);
      return new ReplayOutcome.ExpectedInvalid(
          JazzerHarness.engineCommandSequence().key(),
          expected.getClass().getSimpleName(),
          expected.getMessage(),
          details);
    } catch (IOException unexpected) {
      CommandSequenceDetails details = commandSequenceDetails(input.length, commands);
      return unexpectedFailure(JazzerHarness.engineCommandSequence(), unexpected, details);
    } catch (RuntimeException unexpected) {
      CommandSequenceDetails details = commandSequenceDetails(input.length, commands);
      return unexpectedFailure(JazzerHarness.engineCommandSequence(), unexpected, details);
    }
  }

  private static ReplayOutcome replayRoundTrip(byte[] input) {
    GridGrindFuzzData data = GridGrindFuzzData.replay(input);
    Path directory = null;
    List<WorkbookCommand> commands = List.of();
    try {
      commands = OperationSequenceModel.nextWorkbookCommands(data);
      directory = Files.createTempDirectory("gridgrind-jazzer-replay-");
      Path workbookPath = directory.resolve("workbook.xlsx");
      try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
        new WorkbookCommandExecutor().apply(workbook, commands);
        WorkbookInvariantChecks.requireWorkbookShape(workbook);
        workbook.save(workbookPath);
        XlsxRoundTripVerifier.requireRoundTripReadable(workbook, workbookPath, commands);
      }
      XlsxRoundTripDetails details = roundTripDetails(input.length, commands);
      return new ReplayOutcome.Success(JazzerHarness.xlsxRoundTrip().key(), details);
    } catch (IllegalArgumentException expected) {
      XlsxRoundTripDetails details = roundTripDetails(input.length, commands);
      return new ReplayOutcome.ExpectedInvalid(
          JazzerHarness.xlsxRoundTrip().key(),
          expected.getClass().getSimpleName(),
          expected.getMessage(),
          details);
    } catch (IOException unexpected) {
      XlsxRoundTripDetails details = roundTripDetails(input.length, commands);
      return unexpectedFailure(JazzerHarness.xlsxRoundTrip(), unexpected, details);
    } catch (RuntimeException unexpected) {
      XlsxRoundTripDetails details = roundTripDetails(input.length, commands);
      return unexpectedFailure(JazzerHarness.xlsxRoundTrip(), unexpected, details);
    } finally {
      if (directory != null) {
        deleteRecursively(directory);
      }
    }
  }

  private static ReplayOutcome unexpectedFailure(
      JazzerHarness harness, Throwable error, ReplayDetails details) {
    return new ReplayOutcome.UnexpectedFailure(
        harness.key(),
        error.getClass().getSimpleName(),
        error.getMessage(),
        stackTrace(error),
        details);
  }

  /** Returns the stable external outcome kind used in metadata and reports. */
  static String outcomeKind(ReplayOutcome outcome) {
    Objects.requireNonNull(outcome, "outcome must not be null");
    return switch (outcome) {
      case ReplayOutcome.Success _ -> "SUCCESS";
      case ReplayOutcome.ExpectedInvalid _ -> "EXPECTED_INVALID";
      case ReplayOutcome.UnexpectedFailure _ -> "UNEXPECTED_FAILURE";
    };
  }

  private static ProtocolWorkflowDetails workflowDetails(
      int inputLength, WorkbookPlan request, GridGrindResponse response) {
    if (request == null) {
      return new ProtocolWorkflowDetails(
          inputLength,
          "NOT_GENERATED",
          "NOT_GENERATED",
          0,
          Map.of(),
          Map.of(),
          0,
          Map.of(),
          0,
          Map.of(),
          response == null ? "NOT_EXECUTED" : SequenceIntrospection.responseKind(response));
    }
    return new ProtocolWorkflowDetails(
        inputLength,
        SequenceIntrospection.sourceKind(request),
        SequenceIntrospection.persistenceKind(request),
        request.mutationSteps().size(),
        SequenceIntrospection.mutationKinds(request.mutationSteps()),
        SequenceIntrospection.styleKinds(request.mutationSteps()),
        SequenceIntrospection.assertionCount(request),
        SequenceIntrospection.assertionKinds(request.assertionSteps()),
        SequenceIntrospection.inspectionCount(request),
        SequenceIntrospection.inspectionKinds(request.inspectionSteps()),
        response == null ? "NOT_EXECUTED" : SequenceIntrospection.responseKind(response));
  }

  private static CommandSequenceDetails commandSequenceDetails(
      int inputLength, List<WorkbookCommand> commands) {
    return new CommandSequenceDetails(
        inputLength,
        commands.size(),
        SequenceIntrospection.commandKinds(commands),
        SequenceIntrospection.styleKindsFromCommands(commands));
  }

  private static XlsxRoundTripDetails roundTripDetails(
      int inputLength, List<WorkbookCommand> commands) {
    return new XlsxRoundTripDetails(
        inputLength,
        commands.size(),
        SequenceIntrospection.commandKinds(commands),
        SequenceIntrospection.styleKindsFromCommands(commands));
  }

  private static void deleteRecursively(Path directory) {
    try {
      if (!Files.exists(directory)) {
        return;
      }
      try (var stream = Files.walk(directory)) {
        stream
            .sorted((left, right) -> right.compareTo(left))
            .forEach(
                path -> {
                  try {
                    Files.deleteIfExists(path);
                  } catch (IOException ignored) {
                    // Best-effort cleanup for replay scratch space.
                  }
                });
      }
    } catch (IOException ignored) {
      // Best-effort cleanup for replay scratch space.
    }
  }

  private static String stackTrace(Throwable error) {
    StringWriter stringWriter = new StringWriter();
    try (PrintWriter printWriter = new PrintWriter(stringWriter)) {
      error.printStackTrace(printWriter);
    }
    return stringWriter.toString();
  }
}
