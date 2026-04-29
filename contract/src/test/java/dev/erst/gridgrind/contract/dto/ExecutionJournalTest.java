package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
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
            Optional.of("plan-1"),
            ExecutionJournalLevel.NORMAL,
            new ExecutionJournal.SourceSummary(Optional.of("NEW"), Optional.empty()),
            new ExecutionJournal.PersistenceSummary(Optional.of("NONE"), Optional.empty()),
            ExecutionJournal.Phase.notStarted(),
            ExecutionJournal.Phase.notStarted(),
            ExecutionJournal.Phase.notStarted(),
            new ExecutionJournal.Calculation(
                ExecutionJournal.Phase.notStarted(), ExecutionJournal.Phase.notStarted()),
            ExecutionJournal.Phase.notStarted(),
            ExecutionJournal.Phase.notStarted(),
            List.of(),
            List.of(),
            new ExecutionJournal.Outcome(
                ExecutionJournal.Status.SUCCEEDED,
                0,
                0,
                0,
                Optional.empty(),
                Optional.empty(),
                Optional.empty()),
            List.of());

    assertEquals(ExecutionJournalLevel.NORMAL, journal.level());
    assertEquals(List.of(), journal.warnings());
    assertEquals(List.of(), journal.events());
  }

  @Test
  void phaseRejectsNegativeDurationForStartedStatuses() {
    assertEquals(ExecutionJournal.Status.NOT_STARTED, ExecutionJournal.Phase.notStarted().status());
    assertEquals(
        "NOT_STARTED phases must omit timestamps and use durationMillis=0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Phase(
                        ExecutionJournal.Status.NOT_STARTED,
                        Optional.empty(),
                        Optional.of("2026-04-18T10:00:01Z"),
                        0))
            .getMessage());
    assertEquals(
        "NOT_STARTED phases must omit timestamps and use durationMillis=0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Phase(
                        ExecutionJournal.Status.NOT_STARTED, Optional.empty(), Optional.empty(), 1))
            .getMessage());
    assertEquals(
        "durationMillis must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Phase(
                        ExecutionJournal.Status.SUCCEEDED,
                        Optional.of("2026-04-18T10:00:00Z"),
                        Optional.of("2026-04-18T10:00:01Z"),
                        -1))
            .getMessage());
  }

  @Test
  void stepRejectsNegativeIndexAndInvalidFailureCombinations() {
    ExecutionJournal.Phase phase =
        new ExecutionJournal.Phase(
            ExecutionJournal.Status.SUCCEEDED,
            Optional.of("2026-04-18T10:00:00Z"),
            Optional.of("2026-04-18T10:00:01Z"),
            1);
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
  void calculationRequiresPhases() {
    assertEquals(
        "preflight must not be null",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ExecutionJournal.Calculation(null, ExecutionJournal.Phase.notStarted()))
            .getMessage());
  }

  @Test
  void sourceAndPersistenceSummariesRequireTypeWhenPathIsPresent() {
    assertEquals(
        "type must be present when path is present",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.SourceSummary(
                        Optional.empty(), Optional.of("/tmp/source.xlsx")))
            .getMessage());
    assertEquals(
        "type must be present when path is present",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.PersistenceSummary(
                        Optional.empty(), Optional.of("/tmp/output.xlsx")))
            .getMessage());
  }

  @Test
  void phaseRejectsMissingStartedAndFinishedTimestampsForStartedStatuses() {
    assertEquals(
        "startedAt must be present when status is started",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Phase(
                        ExecutionJournal.Status.SUCCEEDED,
                        Optional.empty(),
                        Optional.of("2026-04-18T10:00:01Z"),
                        1))
            .getMessage());
    assertEquals(
        "finishedAt must be present when status is started",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Phase(
                        ExecutionJournal.Status.SUCCEEDED,
                        Optional.of("2026-04-18T10:00:00Z"),
                        Optional.empty(),
                        1))
            .getMessage());
  }

  @Test
  void outcomeRejectsInvalidCountsDurationsAndFailureOnlyFields() {
    assertEquals(
        "plannedStepCount must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED,
                        -1,
                        0,
                        0,
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "completedStepCount must be >= 0 and <= plannedStepCount",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED,
                        1,
                        -1,
                        0,
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "completedStepCount must be >= 0 and <= plannedStepCount",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED,
                        1,
                        2,
                        0,
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "durationMillis must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED,
                        1,
                        1,
                        -1,
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "failureCode must be present when status is FAILED",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.FAILED,
                        1,
                        0,
                        10,
                        Optional.of(0),
                        Optional.of("step-1"),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "failedStepIndex and failedStepId are only permitted when status is FAILED",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED,
                        1,
                        1,
                        10,
                        Optional.of(0),
                        Optional.empty(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "failedStepIndex and failedStepId are only permitted when status is FAILED",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED,
                        1,
                        1,
                        10,
                        Optional.empty(),
                        Optional.of("step-1"),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "failureCode is only permitted when status is FAILED",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED,
                        1,
                        1,
                        10,
                        Optional.empty(),
                        Optional.empty(),
                        Optional.of(GridGrindProblemCode.INTERNAL_ERROR)))
            .getMessage());
    assertEquals(
        "failedStepIndex must be >= 0 when present",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.FAILED,
                        1,
                        0,
                        10,
                        Optional.of(-1),
                        Optional.of("step-1"),
                        Optional.of(GridGrindProblemCode.INVALID_REQUEST)))
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
