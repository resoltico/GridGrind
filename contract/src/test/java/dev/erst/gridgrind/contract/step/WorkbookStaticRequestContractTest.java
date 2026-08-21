package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.dto.AssertionModeInput;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellRowInput;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Branch coverage for the canonical static request-contract value and rule objects. */
class WorkbookStaticRequestContractTest {
  @Test
  void rejectsInvalidStaticContractValueObjects() {
    assertEquals(
        "index must not be negative",
        assertThrows(
                IllegalArgumentException.class, () -> new WorkbookStaticStep(-1, Optional.empty()))
            .getMessage());
    assertEquals(
        "steps must use contiguous authored indexes",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookStaticRequest(
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        List.of(new WorkbookStaticStep(1, Optional.empty()))))
            .getMessage());
    assertEquals(
        "jsonPath must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> new WorkbookStaticViolation(null, "message"))
            .getMessage());
    assertEquals(
        "jsonPath must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> new WorkbookStaticViolation(" ", "message"))
            .getMessage());
    assertEquals(
        "message must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> new WorkbookStaticViolation("path", null))
            .getMessage());
    assertEquals(
        "message must not be blank",
        assertThrows(IllegalArgumentException.class, () -> new WorkbookStaticViolation("path", " "))
            .getMessage());
  }

  @Test
  void suppressesStreamingSequenceRulesBehindUnboundPredecessors() {
    WorkbookStaticRequest request =
        new WorkbookStaticRequest(
            Optional.empty(),
            Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
            Optional.of(
                ExecutionPolicyInput.modeAndCalculation(
                    ExecutionModeInput.streamingWrite(),
                    CalculationPolicyInput.strategy(new CalculationStrategyInput.EvaluateAll()))),
            List.of(
                new WorkbookStaticStep(0, Optional.empty()),
                new WorkbookStaticStep(1, Optional.of(appendStep("append"))),
                new WorkbookStaticStep(
                    2,
                    Optional.of(
                        new AssertionStep(
                            "assert",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.Text("ready")))))));

    List<String> messages =
        WorkbookStaticRequestContract.validate(request).stream()
            .map(WorkbookStaticViolation::message)
            .toList();

    assertFalse(messages.stream().anyMatch(message -> message.contains("before APPEND_ROW")));
    assertFalse(messages.stream().anyMatch(message -> message.contains("before any assertion")));
    assertFalse(
        messages.stream().anyMatch(message -> message.contains("at least one ENSURE_SHEET")));
  }

  @Test
  void ownsEveryOperationModeBranch() {
    assertTrue(
        CalculationPolicyInput.strategy(new CalculationStrategyInput.ClearCachesOnly())
            .requiresMutationPrefix());
    assertTrue(
        CalculationPolicyInput.strategy(
                new CalculationStrategyInput.EvaluateTargets(
                    List.of(new CellSelector.QualifiedAddress("Ops", "A1"))))
            .requiresMutationPrefix());
    assertTrue(
        WorkbookOperationContracts.executionModeViolation(
                new CellMutationAction.SetCell(new CellInput.NumberValue(1.0)),
                ExecutionModeInput.fullXssf())
            .isEmpty());
    assertTrue(
        WorkbookOperationContracts.executionModeViolation(
                new CellAssertion.CellValue(new CellScalarValue.Text("ready")),
                ExecutionModeInput.streamingWrite())
            .isEmpty());
    assertTrue(
        WorkbookOperationContracts.executionModeViolation(
                new CellAssertion.CellValue(new CellScalarValue.Text("ready")),
                ExecutionModeInput.eventRead())
            .orElseThrow()
            .contains("unsupported step kind: ASSERTION"));
  }

  @Test
  void rejectsExistingWorkbookWritesThatOmitTheTotalSecurityPolicy() {
    WorkbookStaticRequest request =
        new WorkbookStaticRequest(
            Optional.of(new WorkbookPlan.WorkbookSource.ExistingFile("source.xlsx")),
            Optional.of(
                new WorkbookPlan.WorkbookPersistence.SaveAs(
                    "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT)),
            Optional.of(ExecutionPolicyInput.defaults()),
            List.of());

    assertEquals(
        List.of("persistence.security"),
        WorkbookStaticRequestContract.validate(request).stream()
            .map(WorkbookStaticViolation::jsonPath)
            .toList());
  }

  @Test
  void acceptsExistingWorkbookInMemoryRequestsWithoutAWriteSecurityPolicy() {
    WorkbookStaticRequest request =
        new WorkbookStaticRequest(
            Optional.of(new WorkbookPlan.WorkbookSource.ExistingFile("source.xlsx")),
            Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
            Optional.of(ExecutionPolicyInput.defaults()),
            List.of());

    assertTrue(WorkbookStaticPersistenceValidation.validate(request).isEmpty());
  }

  @Test
  void rejectsMutationsAfterTheFirstCollectedAssertion() {
    WorkbookStaticRequest request =
        new WorkbookStaticRequest(
            Optional.of(new WorkbookPlan.WorkbookSource.New()),
            Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
            Optional.of(ExecutionPolicyInput.assertionMode(AssertionModeInput.COLLECT)),
            List.of(
                new WorkbookStaticStep(
                    0,
                    Optional.of(
                        new AssertionStep(
                            "assert",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.Text("ready"))))),
                new WorkbookStaticStep(1, Optional.of(appendStep("append"))),
                new WorkbookStaticStep(2, Optional.of(appendStep("append-2")))));

    assertEquals(
        List.of("steps[1]", "steps[2]"),
        WorkbookStaticRequestContract.validate(request).stream()
            .map(WorkbookStaticViolation::jsonPath)
            .toList());
  }

  @Test
  void preservesCollectedAssertionOrderingAroundAnUnboundPredecessor() {
    WorkbookStaticRequest request =
        new WorkbookStaticRequest(
            Optional.of(new WorkbookPlan.WorkbookSource.New()),
            Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
            Optional.of(ExecutionPolicyInput.assertionMode(AssertionModeInput.COLLECT)),
            List.of(
                new WorkbookStaticStep(0, Optional.empty()),
                new WorkbookStaticStep(
                    1,
                    Optional.of(
                        new AssertionStep(
                            "assert",
                            new CellSelector.ByAddress("Ops", "A1"),
                            new CellAssertion.CellValue(new CellScalarValue.Text("ready"))))),
                new WorkbookStaticStep(2, Optional.of(appendStep("append")))));

    assertEquals(
        List.of("steps[2]"),
        WorkbookStaticAssertionValidation.validate(request).stream()
            .map(WorkbookStaticViolation::jsonPath)
            .toList());
  }

  private static MutationStep appendStep(String stepId) {
    return new MutationStep(
        stepId,
        new SheetSelector.ByName("Ops"),
        new CellMutationAction.AppendRow(
            new CellRowInput.Typed(List.of(new CellInput.Text(TextSourceInput.inline("owner"))))));
  }
}
