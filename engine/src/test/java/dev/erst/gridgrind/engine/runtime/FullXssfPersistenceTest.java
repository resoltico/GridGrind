package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies full-XSSF persistence failures retain or omit publication state truthfully. */
class FullXssfPersistenceTest {
  @TempDir Path root;

  @Test
  void retainsNotPublishedPersistenceFactForAFullXssfPublicationFailure() throws Exception {
    WorkbookResultPersistence.PublicationOutcome.NotPublished publication =
        new WorkbookResultPersistence.PublicationOutcome.NotPublished(
            WorkbookResultPersistence.PublicationOutcome.DestinationState.ABSENT);
    FullXssfPersistence persistence =
        new FullXssfPersistence(
            (workbook, source, requestedPersistence, bindings, acceptance) -> {
              throw new WorkbookPublicationException(
                  publication, new IOException("publication failed"));
            },
            responseSupport());

    try (var workbook = ExcelWorkbooks.create()) {
      FullXssfPersistence.Result result =
          persistence.persist(
              context(workbook), bindings(), CalculationReport.notRequested(), ignored -> {});

      WorkbookResult.Failure failure =
          assertInstanceOf(WorkbookResult.Failure.class, result.failureResponse());
      assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
      assertEquals(
          publication,
          assertInstanceOf(
                  WorkbookResultPersistence.PersistenceOutcome.SavedAs.class, failure.persistence())
              .publication());
    }
  }

  @Test
  void omitsPublicationFactForAnUnexpectedFullXssfPersistenceFailure() throws Exception {
    FullXssfPersistence persistence =
        new FullXssfPersistence(
            (workbook, source, requestedPersistence, bindings, acceptance) -> {
              throw new IOException("storage unavailable");
            },
            responseSupport());

    try (var workbook = ExcelWorkbooks.create()) {
      FullXssfPersistence.Result result =
          persistence.persist(
              context(workbook), bindings(), CalculationReport.notRequested(), ignored -> {});

      WorkbookResult.Failure failure =
          assertInstanceOf(WorkbookResult.Failure.class, result.failureResponse());
      assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    }
  }

  private WorkbookWorkflowExecutionContext context(
      dev.erst.gridgrind.excel.ExcelWorkbook workbook) {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                root.resolve("output.xlsx").toString(),
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    return new WorkbookWorkflowExecutionContext(
        ProtocolVersion.CURRENT,
        request,
        workbook,
        ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
        new ArrayList<>(),
        new ArrayList<>(),
        new ArrayList<>(),
        new ArrayList<>(),
        new ArrayList<>());
  }

  private ExecutionInputBindings bindings() {
    return new ExecutionInputBindings(
        root,
        root.resolve("scratch"),
        new ExecutionGrant(
            List.of(),
            List.of(),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum()));
  }

  private static ExecutionResponseSupport responseSupport() {
    return new ExecutionResponseSupport(
        dev.erst.gridgrind.excel.ExcelWorkbook::close, materialized -> materialized.close());
  }
}
