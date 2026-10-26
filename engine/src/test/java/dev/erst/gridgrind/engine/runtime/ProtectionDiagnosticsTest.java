package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Protection diagnostics omit unrelated formula and coordinate fields. */
class ProtectionDiagnosticsTest {

  @Test
  void executorContextHelpersReturnNoFalseMetadataForWorkbookProtectionOperations() {
    WorkbookMutationAction.SetWorkbookProtection setProtection =
        new WorkbookMutationAction.SetWorkbookProtection(
            new WorkbookProtectionSpec(
                true,
                false,
                true,
                Optional.of(new dev.erst.gridgrind.core.source.SecretReference("book-password")),
                Optional.empty()));
    WorkbookMutationAction.ClearWorkbookProtection clearProtection =
        new WorkbookMutationAction.ClearWorkbookProtection();

    assertEquals("SET_WORKBOOK_PROTECTION", setProtection.actionType());
    assertEquals("CLEAR_WORKBOOK_PROTECTION", clearProtection.actionType());
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.formulaFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.sheetNameFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.addressFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.rangeFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.namedRangeNameFor(
            new MutationStep(
                new StepId("step-01-set-workbook-protection"),
                new WorkbookSelector.Current(),
                setProtection),
            new IllegalStateException()));

    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.formulaFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.sheetNameFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.addressFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.rangeFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionDiagnosticFields.namedRangeNameFor(
            new MutationStep(
                new StepId("step-02-clear-workbook-protection"),
                new WorkbookSelector.Current(),
                clearProtection),
            new IllegalStateException()));
  }
}
