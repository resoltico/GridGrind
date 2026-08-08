package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct validation coverage for the execution-journal contract family. */
class ExecutionJournalTest {
  @Test
  void defaultsLevelWarningsAndEvents() {
    ExecutionJournal journal =
        new ExecutionJournal(
            ExecutionJournalLevel.NORMAL,
            new ExecutionJournal.SourceSummary(Optional.of("NEW"), Optional.empty()),
            ExecutionJournal.Phase.notStarted(),
            ExecutionJournal.Phase.notStarted(),
            ExecutionJournal.Phase.notStarted(),
            new ExecutionJournal.Calculation(
                ExecutionJournal.Phase.notStarted(), ExecutionJournal.Phase.notStarted()),
            ExecutionJournal.Phase.notStarted(),
            ExecutionJournal.Phase.notStarted(),
            List.of(),
            ExecutionJournal.Outcome.succeeded(0, 0, 0),
            List.of());

    assertEquals(ExecutionJournalLevel.NORMAL, journal.level());
    assertEquals(List.of(), journal.events());
  }

  @Test
  void phaseFactoriesReturnTypedVariantsAndValidateTiming() {
    assertEquals(ExecutionJournal.Status.NOT_STARTED, ExecutionJournal.Phase.notStarted().status());
    assertEquals(
        ExecutionJournal.Status.NOT_REQUESTED, ExecutionJournal.Phase.notRequested().status());
    ExecutionJournal.Phase.Succeeded succeeded =
        assertInstanceOf(
            ExecutionJournal.Phase.Succeeded.class,
            ExecutionJournal.Phase.succeeded("2026-04-18T10:00:00Z", "2026-04-18T10:00:01Z", 1));
    assertEquals("2026-04-18T10:00:00Z", succeeded.timing().orElseThrow().startedAt());
    assertEquals(1L, succeeded.timing().orElseThrow().durationMillis());
    assertEquals(
        "durationMillis must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Timing("2026-04-18T10:00:00Z", "2026-04-18T10:00:01Z", -1))
            .getMessage());
  }

  @Test
  void stepRejectsNegativeIndexAndInvalidFailureCombinations() {
    ExecutionJournal.Phase phase =
        ExecutionJournal.Phase.succeeded("2026-04-18T10:00:00Z", "2026-04-18T10:00:01Z", 1);
    ExecutionJournal.FailureClassification failure =
        new ExecutionJournal.FailureClassification(
            GridGrindProblemCode.ASSERTION_FAILED,
            GridGrindProblemCategory.REQUEST,
            "EXECUTE_STEP",
            "boom");

    assertEquals(
        "stepIndex must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Step(
                        -1,
                        "step-1",
                        "ASSERTION",
                        "EXPECT_CELL_VALUE",
                        List.of(new ExecutionJournal.Target("CELL", "Cell Budget!B4")),
                        phase,
                        ExecutionJournal.StepOutcome.SUCCEEDED,
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "failure must be present when outcome is FAILED",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Step(
                        0,
                        "step-1",
                        "ASSERTION",
                        "EXPECT_CELL_VALUE",
                        List.of(new ExecutionJournal.Target("CELL", "Cell Budget!B4")),
                        phase,
                        ExecutionJournal.StepOutcome.FAILED,
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "failure is only permitted when outcome is FAILED",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Step(
                        0,
                        "step-1",
                        "ASSERTION",
                        "EXPECT_CELL_VALUE",
                        List.of(new ExecutionJournal.Target("CELL", "Cell Budget!B4")),
                        phase,
                        ExecutionJournal.StepOutcome.SUCCEEDED,
                        Optional.of(failure)))
            .getMessage());
  }

  @Test
  void outcomeVariantsValidateCountsAndFailureShape() {
    ExecutionJournal.Outcome.Succeeded succeeded =
        assertInstanceOf(
            ExecutionJournal.Outcome.Succeeded.class, ExecutionJournal.Outcome.succeeded(2, 2, 10));
    assertEquals(2, succeeded.plannedStepCount());
    assertEquals(10L, succeeded.durationMillis());
    ExecutionJournal.Outcome.Failed failed =
        assertInstanceOf(
            ExecutionJournal.Outcome.Failed.class,
            ExecutionJournal.Outcome.failed(
                2,
                1,
                10,
                GridGrindProblemCode.INVALID_REQUEST,
                Optional.of(new ExecutionJournal.FailureStep(1, "step-2"))));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failed.problemCode());
    assertEquals("step-2", failed.failedStep().orElseThrow().failedStepId());
    assertEquals(
        "plannedStepCount must be >= 0",
        assertThrows(
                IllegalArgumentException.class, () -> ExecutionJournal.Outcome.succeeded(-1, 0, 0))
            .getMessage());
    assertEquals(
        "completedStepCount must be >= 0 and <= plannedStepCount",
        assertThrows(
                IllegalArgumentException.class, () -> ExecutionJournal.Outcome.succeeded(1, 2, 0))
            .getMessage());
    assertEquals(
        "completedStepCount must be >= 0 and <= plannedStepCount",
        assertThrows(
                IllegalArgumentException.class, () -> ExecutionJournal.Outcome.succeeded(1, -1, 0))
            .getMessage());
    assertEquals(
        "failedStepIndex must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ExecutionJournal.FailureStep(-1, "step-1"))
            .getMessage());
  }

  @Test
  void calculationRequiresPhases() {
    assertEquals(
        "preflight must not be null",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ExecutionJournal.Calculation(null, ExecutionJournal.Phase.notStarted()))
            .getMessage());
  }

  @Test
  void sourceSummariesRequireTypeWhenPathIsPresent() {
    assertEquals(
        "type must be present when path is present",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.SourceSummary(
                        Optional.empty(), Optional.of("/tmp/source.xlsx")))
            .getMessage());
  }

  @Test
  void eventRequiresStepIndexAndStepIdTogether() {
    assertEquals(
        "stepId and stepIndex must either both be present or both be absent",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Event(
                        "2026-04-18T10:00:00Z",
                        "STEP",
                        "Started",
                        Optional.of(1),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "stepId and stepIndex must either both be present or both be absent",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Event(
                        "2026-04-18T10:00:00Z",
                        "STEP",
                        "Started",
                        Optional.empty(),
                        Optional.of("step-1")))
            .getMessage());
  }
}
