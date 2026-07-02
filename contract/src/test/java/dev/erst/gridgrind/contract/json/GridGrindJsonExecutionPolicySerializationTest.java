package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
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
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {},
                      "steps": []
                    }
                    """));

    assertEquals(ExecutionModeInput.defaults(), request.execution().mode());
    assertEquals(ExecutionJournalInput.defaults(), request.execution().journal());
    assertEquals(CalculationPolicyInput.defaults(), request.execution().calculation());
  }

  @Test
  void requestAcceptsModeOnlyExecutionBlocksAndDefaultsRemainingAxes() {
    WorkbookPlan request =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "mode": { "type": "EVENT_READ" }
                      },
                      "steps": []
                    }
                    """));

    assertEquals(ExecutionModeInput.eventRead(), request.execution().mode());
    assertEquals(ExecutionJournalInput.defaults(), request.execution().journal());
    assertEquals(CalculationPolicyInput.defaults(), request.execution().calculation());
  }

  @Test
  void requestAcceptsJournalOnlyExecutionBlocksAndDefaultsOtherAxes() {
    WorkbookPlan request =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "journal": { "level": "VERBOSE" }
                      },
                      "steps": []
                    }
                    """));

    assertEquals(ExecutionModeInput.defaults(), request.execution().mode());
    assertEquals(ExecutionJournalLevel.VERBOSE, request.execution().journal().level());
    assertEquals(CalculationPolicyInput.defaults(), request.execution().calculation());
  }

  @Test
  void requestAcceptsCalculationOnlyExecutionBlocksAndDefaultsRemainingAxes() {
    WorkbookPlan request =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V1",
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

    assertEquals(ExecutionModeInput.defaults(), request.execution().mode());
    assertEquals(ExecutionJournalInput.defaults(), request.execution().journal());
    assertEquals(
        new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true),
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
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "execution": {
                        "calculation": {}
                      },
                      "steps": []
                    }
                    """));

    assertEquals(ExecutionModeInput.defaults(), request.execution().mode());
    assertEquals(ExecutionJournalInput.defaults(), request.execution().journal());
    assertEquals(CalculationPolicyInput.defaults(), request.execution().calculation());
  }

  @Test
  void requestSerializationOmitsDefaultExecutionAxesAndNestedCalculationFields() {
    WorkbookPlan modeOnlyRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.mode(ExecutionModeInput.eventRead()),
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan journalOnlyRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.journal(new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan calculationOnlyRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.calculation(
                new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true)),
            FormulaEnvironmentInput.empty(),
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
