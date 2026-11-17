package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionJournalSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.PresenceAssertion;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.GridGrindWarningCode;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.result.ExecutionProgressEvent;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/** Optional progress delivery cannot rewrite execution facts or publication occurrence. */
class ProgressFailureIsolationTest {
  @TempDir Path root;

  @ParameterizedTest
  @CsvSource({
    "PLAN,STARTED",
    "STEP,SUCCEEDED",
    "PERSIST,STARTED",
    "PERSIST,SUCCEEDED",
    "CLOSE,SUCCEEDED"
  })
  void observerRuntimeFailureDisablesDeliveryOnceAndRetainsCompleteSuccess(
      ExecutionProgressEvent.Category category, ExecutionProgressEvent.Status status)
      throws Exception {
    AtomicInteger failures = new AtomicInteger();
    AtomicInteger callbacksAfterFailure = new AtomicInteger();
    var result =
        GridGrindEngine.requestExecutor()
            .execute(
                plan(),
                inputs(),
                event -> {
                  if (failures.get() > 0) callbacksAfterFailure.incrementAndGet();
                  if (event.category() == category && event.status() == status) {
                    failures.incrementAndGet();
                    throw new IllegalStateException("synthetic observer failure");
                  }
                });
    WorkbookResult.Success success = assertInstanceOf(WorkbookResult.Success.class, result);
    assertEquals(1, failures.get());
    assertEquals(0, callbacksAfterFailure.get());
    assertEquals(
        1,
        success.warnings().stream()
            .filter(warning -> warning.code() == GridGrindWarningCode.PROGRESS_SINK_FAILED)
            .count());
    assertEquals(1, success.assertions().size());
    assertEquals(1, success.inspections().size());
    var persistence = assertInstanceOf(PersistenceOutcome.SavedAs.class, success.persistence());
    assertInstanceOf(PublicationOutcome.Published.class, persistence.publication());
    assertTrue(Files.size(root.resolve("output.xlsx")) > 0);
    try (var workbook = new XSSFWorkbook(root.resolve("output.xlsx").toFile())) {
      assertEquals("Budget", workbook.getSheetName(0));
    }
  }

  @Test
  void callerErrorEscapesWithoutBeingConvertedToAWorkbookFailure() {
    AssertionError sentinel = new AssertionError("synthetic caller error");
    assertSame(
        sentinel,
        assertThrows(
            AssertionError.class,
            () ->
                GridGrindEngine.requestExecutor()
                    .execute(
                        plan(),
                        inputs(),
                        event -> {
                          throw sentinel;
                        })));
  }

  private WorkbookPlan plan() {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
        ExecutionPolicySpec.journal(new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE)),
        FormulaEnvironmentSpec.empty(),
        List.of(
            new MutationStep(
                new StepId("ensure"),
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.EnsureSheet()),
            new AssertionStep(
                new StepId("verify"),
                new SheetSelector.ByName(new SheetName("Budget")),
                new PresenceAssertion.SheetPresent()),
            new InspectionStep(
                new StepId("inspect"),
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary())));
  }

  private GridGrindRequestInputs inputs() {
    return new GridGrindRequestInputs(
        root,
        root.resolve("scratch"),
        new ExecutionGrant(
            0,
            List.of(),
            List.of("ENSURE_SHEET", "EXPECT_SHEET_PRESENT", "GET_WORKBOOK_SUMMARY"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.SaveAs(
                root.resolve("output.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            List.of(),
            AcceptancePolicy.minimum()));
  }
}
