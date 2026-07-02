package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
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

    GridGrindResponsePersistence.PersistenceOutcome.Overwritten overwritten =
        assertInstanceOf(
            GridGrindResponsePersistence.PersistenceOutcome.Overwritten.class,
            ExecutionRequestPaths.unwrittenPersistenceOutcome(overwriteExistingRequest));
    GridGrindResponsePersistence.PersistenceOutcome.SavedAs savedAs =
        assertInstanceOf(
            GridGrindResponsePersistence.PersistenceOutcome.SavedAs.class,
            ExecutionRequestPaths.unwrittenPersistenceOutcome(saveAsRequest));
    GridGrindResponsePersistence.PersistenceOutcome.Overwritten impossibleOverwrite =
        assertInstanceOf(
            GridGrindResponsePersistence.PersistenceOutcome.Overwritten.class,
            ExecutionRequestPaths.unwrittenPersistenceOutcome(impossibleOverwriteRequest));

    assertEquals(java.util.Optional.of("fixtures/budget.xlsx"), overwritten.sourcePath());
    assertInstanceOf(
        GridGrindResponsePersistence.WriteResult.NotWritten.class, overwritten.write());
    assertEquals("fixtures/output.xlsx", savedAs.requestedPath());
    assertInstanceOf(GridGrindResponsePersistence.WriteResult.NotWritten.class, savedAs.write());
    assertEquals(java.util.Optional.empty(), impossibleOverwrite.sourcePath());
    assertInstanceOf(
        GridGrindResponsePersistence.WriteResult.NotWritten.class, impossibleOverwrite.write());
  }
}
