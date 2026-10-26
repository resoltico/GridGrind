package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.CalculationStrategySpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionJournalSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;

/** Locks independent execution-policy axis defaulting at the request wire boundary. */
final class GridGrindJsonExecutionPolicySerializationTest {
  @Test
  void requestAcceptsEmptyExecutionBlocksAndDefaultsEveryAxis() {
    WorkbookPlan request =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": 1,
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {},
                      "steps": []
                    }
                    """));

    assertEquals(ExecutionModeSpec.defaults(), request.execution().mode());
    assertEquals(ExecutionJournalSpec.defaults(), request.execution().journal());
    assertEquals(CalculationPolicySpec.defaults(), request.execution().calculation());
  }

  @Test
  void requestAcceptsModeOnlyExecutionBlocksAndDefaultsRemainingAxes() {
    WorkbookPlan request =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": 1,
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": { "type": "EVENT_READ" }
                      },
                      "steps": []
                    }
                    """));

    assertEquals(ExecutionModeSpec.eventRead(), request.execution().mode());
    assertEquals(ExecutionJournalSpec.defaults(), request.execution().journal());
    assertEquals(CalculationPolicySpec.defaults(), request.execution().calculation());
  }

  @Test
  void requestAcceptsJournalOnlyExecutionBlocksAndDefaultsOtherAxes() {
    WorkbookPlan request =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": 1,
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "journal": { "level": "VERBOSE" }
                      },
                      "steps": []
                    }
                    """));

    assertEquals(ExecutionModeSpec.defaults(), request.execution().mode());
    assertEquals(ExecutionJournalLevel.VERBOSE, request.execution().journal().level());
    assertEquals(CalculationPolicySpec.defaults(), request.execution().calculation());
  }

  @Test
  void requestAcceptsCalculationOnlyExecutionBlocksAndDefaultsRemainingAxes() {
    WorkbookPlan request =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": 1,
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "calculation": {
                          "markRecalculateOnOpen": true
                        }
                      },
                      "steps": []
                    }
                    """));

    assertEquals(ExecutionModeSpec.defaults(), request.execution().mode());
    assertEquals(ExecutionJournalSpec.defaults(), request.execution().journal());
    assertEquals(
        new CalculationPolicySpec(new CalculationStrategySpec.DoNotCalculate(), true),
        request.execution().calculation());
  }

  @Test
  void requestAcceptsEmptyCalculationBlocksAndDefaultsNestedCalculationFields() {
    WorkbookPlan request =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": 1,
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "calculation": {}
                      },
                      "steps": []
                    }
                    """));

    assertEquals(ExecutionModeSpec.defaults(), request.execution().mode());
    assertEquals(ExecutionJournalSpec.defaults(), request.execution().journal());
    assertEquals(CalculationPolicySpec.defaults(), request.execution().calculation());
  }

  @Test
  void requestSerializationOmitsDefaultExecutionAxesAndNestedCalculationFields() {
    WorkbookPlan modeOnlyRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.mode(ExecutionModeSpec.eventRead()),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookPlan journalOnlyRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.journal(new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE)),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookPlan calculationOnlyRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.calculation(
                new CalculationPolicySpec(new CalculationStrategySpec.DoNotCalculate(), true)),
            FormulaEnvironmentSpec.empty(),
            List.of());

    JsonNode modeExecution = GridGrindJsonOutput.requestTree(modeOnlyRequest).path("execution");
    JsonNode journalExecution =
        GridGrindJsonOutput.requestTree(journalOnlyRequest).path("execution");
    JsonNode calculationExecution =
        GridGrindJsonOutput.requestTree(calculationOnlyRequest).path("execution");

    assertEquals("EVENT_READ", modeExecution.path("mode").path("type").stringValue());
    assertFalse(modeExecution.has("journal"));
    assertFalse(modeExecution.has("calculation"));

    assertEquals("VERBOSE", journalExecution.path("journal").path("level").stringValue());
    assertFalse(journalExecution.has("mode"));
    assertFalse(journalExecution.has("calculation"));

    assertTrue(
        calculationExecution.path("calculation").path("markRecalculateOnOpen").booleanValue());
    assertFalse(calculationExecution.has("mode"));
    assertFalse(calculationExecution.has("journal"));
    assertFalse(calculationExecution.path("calculation").has("strategy"));
  }
}
