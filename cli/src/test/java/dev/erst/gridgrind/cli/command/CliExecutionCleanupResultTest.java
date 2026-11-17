package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.authoring.GridGrindPlan;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.core.result.WorkbookResults;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import org.junit.jupiter.api.Test;

/** Verifies post-execution cleanup cannot erase publication or completed observation facts. */
class CliExecutionCleanupResultTest {
  @org.junit.jupiter.api.io.TempDir java.nio.file.Path root;

  @Test
  void executorSideEffectFollowedByUnexpectedFailureIsNeverReportedAsNotAttempted()
      throws Exception {
    var output = root.resolve("output.xlsx");
    var request =
        GridGrindPlan.newWorkbook()
            .saveAs(
                output,
                dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence.IfExists.REJECT)
            .toPlan();
    var failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            CliExecutionFailureSupport.executeBound(
                request,
                root,
                () -> {},
                () -> {
                  try {
                    java.nio.file.Files.writeString(output, "custom executor side effect");
                  } catch (IOException exception) {
                    throw new java.io.UncheckedIOException(exception);
                  }
                  throw new IllegalStateException(
                      "custom executor failed without returning its facts");
                }));
    var saved =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.SavedAs.class, failure.persistence());
    var uncertain =
        assertInstanceOf(
            WorkbookResultPersistence.PublicationOutcome.Uncertain.class, saved.publication());
    assertEquals(output.toString(), uncertain.executionPath());
    assertEquals(
        WorkbookResultPersistence.PublicationOutcome.CommitObservation.NOT_ESTABLISHED,
        uncertain.commitObservation());
    assertEquals("custom executor side effect", java.nio.file.Files.readString(output));
  }

  @Test
  void postResultCloseFailureRetainsPublicationAndCompletedObservationFacts() {
    var publication =
        new WorkbookResultPersistence.PublicationOutcome.Published(
            "/synthetic/output.xlsx",
            "0".repeat(64),
            12,
            new WorkbookResultPersistence.PublicationOutcome.StagedArtifactVerification(),
            new WorkbookResultPersistence.PublicationOutcome.DurabilityEvidence
                .FileAndDirectorySynced());
    var prior =
        WorkbookResults.success(
            ProtocolVersion.CURRENT,
            new WorkbookResultPersistence.PersistenceOutcome.SavedAs("output.xlsx", publication),
            List.of(),
            List.of(new AssertionResult.Passed(new StepId("checked"), "SHEET_PRESENT")),
            List.of(
                new WorkbookInspectionResult.PackageSecurityResult(
                    new StepId("observed"), OoxmlPackageSecurityReport.none())));
    var request = GridGrindPlan.newWorkbook().toPlan();
    for (Exception cleanup :
        List.of(
            new IOException("synthetic cleanup IO"),
            new IllegalStateException("synthetic cleanup runtime"))) {
      var result =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              CliExecutionFailureSupport.executeBound(
                  request,
                  java.nio.file.Path.of("/synthetic"),
                  () -> {
                    throw cleanup;
                  },
                  () -> prior));
      assertEquals(prior.persistence(), result.persistence());
      assertEquals(prior.evidence(), result.evidence());
      assertEquals(prior.assertions(), result.assertions());
      assertEquals(prior.inspections(), result.inspections());
      assertInstanceOf(
          dev.erst.gridgrind.core.result.ExecutionJournal.Outcome.Failed.class,
          result.journal().outcome());
    }
  }

  @Test
  void executorErrorPropagatesAfterBindingCloseAndRetainsCleanupAsSecondary() {
    var marker = new AssertionError("synthetic executor Error");
    var cleanup = new IOException("synthetic binding close failure");
    var closed = new AtomicBoolean();
    assertSame(
        marker,
        assertThrows(
            AssertionError.class,
            () ->
                CliExecutionFailureSupport.executeBound(
                    GridGrindPlan.newWorkbook().toPlan(),
                    java.nio.file.Path.of("/synthetic"),
                    () -> {
                      closed.set(true);
                      throw cleanup;
                    },
                    () -> {
                      throw marker;
                    })));
    assertTrue(closed.get());
    assertEquals(1, marker.getSuppressed().length);
    assertSame(cleanup, marker.getSuppressed()[0]);
  }
}
