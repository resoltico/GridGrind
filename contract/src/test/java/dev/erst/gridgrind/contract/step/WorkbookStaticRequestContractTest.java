package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.AssertionModeInput;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellRowInput;
import dev.erst.gridgrind.contract.dto.CellScalarValue;
import dev.erst.gridgrind.contract.dto.CellStylePatchInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
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
  void rejectsReversedRangesBeforeWorkbookExecution() {
    assertEquals("A1", new RangeSelector.ByRange("Ops", "A1").range());
    assertEquals(
        "range must end at or below and to the right of its start address",
        assertThrows(
                dev.erst.gridgrind.contract.json.InvalidRequestException.class,
                () -> new RangeSelector.ByRange("Ops", "B5:A1"))
            .getMessage());
    assertEquals(
        "range must end at or below and to the right of its start address",
        assertThrows(
                dev.erst.gridgrind.contract.json.InvalidRequestException.class,
                () -> new RangeSelector.ByRange("Ops", "B1:A5"))
            .getMessage());
  }

  @Test
  void rejectsRangeMaterializationBeyondThePublishedPlanBudget() {
    MutationStep styleWholeColumn =
        new MutationStep(
            "style",
            new RangeSelector.ByRange("Ops", "A1:A250001"),
            new CellMutationAction.ClearRange());

    WorkbookStaticViolation violation =
        WorkbookStaticRequestContract.validate(staticRequest(List.of(styleWholeColumn))).getFirst();

    assertEquals("steps[0].target.range", violation.jsonPath());
    assertEquals(
        "materialized worksheet work must not exceed 250000 items per plan; this step raises the total to 250001",
        violation.message());
  }

  @Test
  void rejectsSetRangeWhoseTargetAndAuthoredGridHaveDifferentDimensions() {
    MutationStep setRange =
        new MutationStep(
            "range",
            new RangeSelector.ByRange("Ops", "A1:C3"),
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.contract.dto.CellGridInput.NumberRows(
                    List.of(List.of(1.0d)))));

    WorkbookStaticViolation violation =
        WorkbookStaticRequestContract.validate(staticRequest(List.of(setRange))).getFirst();

    assertEquals("steps[0].action.rows", violation.jsonPath());
    assertEquals(
        "range dimensions do not match provided values: A1:C3 expects 9 cells but received 1",
        violation.message());
  }

  @Test
  void acceptsMatchingSmallRangesAndRejectsCumulativeMaterializationOverflow() {
    MutationStep matchingRange =
        new MutationStep(
            "matching",
            new RangeSelector.ByRange("Ops", "A1:B1"),
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.contract.dto.CellGridInput.NumberRows(
                    List.of(List.of(1.0d, 2.0d)))));
    MutationStep firstLargeRange =
        new MutationStep(
            "first",
            new RangeSelector.ByRange("Ops", "A1:A125000"),
            new CellMutationAction.ClearRange());
    MutationStep secondLargeRange =
        new MutationStep(
            "second",
            new RangeSelector.ByRange("Ops", "B1:B125001"),
            new CellMutationAction.ClearRange());
    MutationStep arrayFormula =
        new MutationStep(
            "array-formula",
            new RangeSelector.ByRange("Ops", "C1:C2"),
            new CellMutationAction.SetArrayFormula(
                new ArrayFormulaInput(TextSourceInput.inline("SUM(A1:A2)"))));
    MutationStep styledRange =
        new MutationStep(
            "styled-range",
            new RangeSelector.ByRange("Ops", "D1:D2"),
            new CellMutationAction.ApplyStyle(
                new CellStylePatchInput(
                    Optional.of("0.00"),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())));
    MutationStep singleCell =
        new MutationStep(
            "single-cell",
            new RangeSelector.ByRange("Ops", "E1"),
            new CellMutationAction.SetCell(new CellInput.NumberValue(1.0d)));
    MutationStep workbookAction =
        new MutationStep(
            "workbook-action",
            new RangeSelector.ByRange("Ops", "F1"),
            new WorkbookMutationAction.EnsureSheet());

    assertTrue(
        WorkbookStaticRequestContract.validate(
                staticRequest(List.of(matchingRange, arrayFormula, styledRange)))
            .isEmpty());
    assertEquals(
        List.of("steps[3].target.type", "steps[4].target.type"),
        WorkbookStaticRequestContract.validate(
                staticRequest(
                    List.of(matchingRange, arrayFormula, styledRange, singleCell, workbookAction)))
            .stream()
            .map(WorkbookStaticViolation::jsonPath)
            .toList());
    assertEquals(
        List.of("steps[2].target.range"),
        WorkbookStaticRequestContract.validate(
                staticRequest(List.of(matchingRange, firstLargeRange, secondLargeRange)))
            .stream()
            .map(WorkbookStaticViolation::jsonPath)
            .toList());
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
                    List.of(new dev.erst.gridgrind.contract.dto.FormulaCellTarget("Ops", "A1"))))
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

  @Test
  void rejectsOnlyProvableFormulaThenColumnEditOrderings() {
    MutationStep formula =
        new MutationStep(
            "formula",
            new CellSelector.ByAddress("Ops", "A1"),
            new CellMutationAction.SetCell(new CellInput.Formula(TextSourceInput.inline("1+1"))));
    MutationStep columnEdit =
        new MutationStep(
            "insert-column",
            new ColumnBandSelector.Insertion("Ops", 1, 1),
            new WorkbookMutationAction.InsertColumns());
    WorkbookStaticRequest formulaThenColumn = staticRequest(List.of(formula, columnEdit));

    assertEquals(
        List.of("steps[1].action"),
        WorkbookStaticFormulaColumnValidation.validate(formulaThenColumn).stream()
            .map(WorkbookStaticViolation::jsonPath)
            .toList());
    assertEquals(
        List.of(WorkbookStaticFormulaColumnValidation.VIOLATION_MESSAGE),
        WorkbookStaticFormulaColumnValidation.validate(formulaThenColumn).stream()
            .map(WorkbookStaticViolation::message)
            .toList());

    assertTrue(
        WorkbookStaticFormulaColumnValidation.validate(staticRequest(List.of(columnEdit, formula)))
            .isEmpty());
    assertTrue(
        WorkbookStaticFormulaColumnValidation.validate(
                staticRequest(
                    List.of(
                        formula,
                        new MutationStep(
                            "replace",
                            new RangeSelector.ByRange("Ops", "A1"),
                            new CellMutationAction.ClearRange()),
                        columnEdit)))
            .isEmpty());
    MutationStep rawFormula =
        new MutationStep(
            "raw-formula",
            new CellSelector.ByAddress("Ops", "A2"),
            new CellMutationAction.SetCell(
                new CellInput.RawFormula(TextSourceInput.inline("RAW_FORMULA()"))));
    assertEquals(
        List.of("steps[1].action"),
        WorkbookStaticFormulaColumnValidation.validate(
                staticRequest(List.of(rawFormula, columnEdit)))
            .stream()
            .map(WorkbookStaticViolation::jsonPath)
            .toList());
  }

  private static WorkbookStaticRequest staticRequest(List<WorkbookStep> steps) {
    return new WorkbookStaticRequest(
        Optional.of(new WorkbookPlan.WorkbookSource.New()),
        Optional.of(new WorkbookPlan.WorkbookPersistence.None()),
        Optional.of(ExecutionPolicyInput.defaults()),
        java.util.stream.IntStream.range(0, steps.size())
            .mapToObj(index -> new WorkbookStaticStep(index, Optional.of(steps.get(index))))
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
