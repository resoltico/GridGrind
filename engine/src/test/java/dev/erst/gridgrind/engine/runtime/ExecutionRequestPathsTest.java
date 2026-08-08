package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Contract tests for request-derived path and persistence facts. */
class ExecutionRequestPathsTest {
  @Test
  void unwrittenPersistenceOutcomesPreserveSaveIntentWithoutInventingPaths() {
    WorkbookPlan overwriteExistingRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("fixtures/budget.xlsx"),
            new WorkbookPlan.WorkbookPersistence.Overwrite(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan saveAsRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "fixtures/output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan impossibleOverwriteRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.Overwrite(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());

    WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
            ExecutionRequestPaths.unwrittenPersistenceOutcome(overwriteExistingRequest));
    WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.SavedAs.class,
            ExecutionRequestPaths.unwrittenPersistenceOutcome(saveAsRequest));
    WorkbookResultPersistence.PersistenceOutcome.Overwritten impossibleOverwrite =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
            ExecutionRequestPaths.unwrittenPersistenceOutcome(impossibleOverwriteRequest));

    assertEquals(java.util.Optional.of("fixtures/budget.xlsx"), overwritten.sourcePath());
    assertInstanceOf(WorkbookResultPersistence.WriteResult.NotWritten.class, overwritten.write());
    assertEquals("fixtures/output.xlsx", savedAs.requestedPath());
    assertInstanceOf(WorkbookResultPersistence.WriteResult.NotWritten.class, savedAs.write());
    assertEquals(java.util.Optional.empty(), impossibleOverwrite.sourcePath());
    assertInstanceOf(
        WorkbookResultPersistence.WriteResult.NotWritten.class, impossibleOverwrite.write());
  }
}
