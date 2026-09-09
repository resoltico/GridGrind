package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
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
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    return new WorkbookWorkflowExecutionContext(
        GridGrindProtocolVersion.current(),
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
        new GridGrindExecutionGrant.Bounded(
            List.of(),
            List.of(),
            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new GridGrindExecutionGrant.PublicationAuthority.None(),
            List.of(),
            GridGrindHostAcceptancePolicy.minimum()));
  }

  private static ExecutionResponseSupport responseSupport() {
    return new ExecutionResponseSupport(
        dev.erst.gridgrind.excel.ExcelWorkbook::close, materialized -> materialized.close());
  }
}
