package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct validation coverage for the execution-journal contract family. */
class ExecutionJournalTest {
  @Test
  void defaultsLevelWarningsAndEvents() {
    ExecutionJournal journal =
        new ExecutionJournal(
            "plan-1",
            null,
            new ExecutionJournal.SourceSummary("NEW", null),
            new ExecutionJournal.PersistenceSummary("NONE", null),
            ExecutionJournal.Phase.notStarted(),
            ExecutionJournal.Phase.notStarted(),
            ExecutionJournal.Phase.notStarted(),
            new ExecutionJournal.Calculation(
                ExecutionJournal.Phase.notStarted(), ExecutionJournal.Phase.notStarted()),
            ExecutionJournal.Phase.notStarted(),
            ExecutionJournal.Phase.notStarted(),
            List.of(),
            null,
            new ExecutionJournal.Outcome(
                ExecutionJournal.Status.SUCCEEDED, 0, 0, 0, null, null, null),
            null);

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
                        ExecutionJournal.Status.NOT_STARTED, null, "2026-04-18T10:00:01Z", 0))
            .getMessage());
    assertEquals(
        "NOT_STARTED phases must omit timestamps and use durationMillis=0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Phase(ExecutionJournal.Status.NOT_STARTED, null, null, 1))
            .getMessage());
    assertEquals(
        "durationMillis must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Phase(
                        ExecutionJournal.Status.SUCCEEDED,
                        "2026-04-18T10:00:00Z",
                        "2026-04-18T10:00:01Z",
                        -1))
            .getMessage());
  }

  @Test
  void stepRejectsNegativeIndexAndInvalidFailureCombinations() {
    ExecutionJournal.Phase phase =
        new ExecutionJournal.Phase(
            ExecutionJournal.Status.SUCCEEDED, "2026-04-18T10:00:00Z", "2026-04-18T10:00:01Z", 1);
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
                        null))
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
                        null))
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
                        failure))
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
  void outcomeRejectsInvalidCountsDurationsAndFailureOnlyFields() {
    assertEquals(
        "plannedStepCount must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED, -1, 0, 0, null, null, null))
            .getMessage());
    assertEquals(
        "completedStepCount must be >= 0 and <= plannedStepCount",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED, 1, -1, 0, null, null, null))
            .getMessage());
    assertEquals(
        "completedStepCount must be >= 0 and <= plannedStepCount",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED, 1, 2, 0, null, null, null))
            .getMessage());
    assertEquals(
        "durationMillis must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED, 1, 1, -1, null, null, null))
            .getMessage());
    assertEquals(
        "failureCode must be present when status is FAILED",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.FAILED, 1, 0, 10, 0, "step-1", null))
            .getMessage());
    assertEquals(
        "failedStepIndex and failedStepId are only permitted when status is FAILED",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED, 1, 1, 10, 0, null, null))
            .getMessage());
    assertEquals(
        "failedStepIndex and failedStepId are only permitted when status is FAILED",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Outcome(
                        ExecutionJournal.Status.SUCCEEDED, 1, 1, 10, null, "step-1", null))
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
                        null,
                        null,
                        GridGrindProblemCode.INTERNAL_ERROR))
            .getMessage());
  }
}
