package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
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
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.result.ExecutionProgressEvent;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResults;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Independent counterexamples for facts and cleanup challenged in the foundation review. */
class FoundationFailureEvidenceTest {
  @TempDir Path root;

  @Test
  void failedPublicationRetainsAdmissionCompletedObservationsAndVerifiedStage() throws Exception {
    Files.writeString(root.resolve("label.txt"), "Label");
    Path output = root.resolve("output.xlsx");
    WorkbookPlan request =
        plan(
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            List.of(
                ensureSheet(),
                setText("write", "A1", new TextSourceSpec.Utf8File("label.txt")),
                sheetPresent("check", "Budget"),
                new InspectionStep(
                    new StepId("observe"),
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    var bindings =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            ExecutionGrantTestSupport.forPlan(request, root, root.resolve("label.txt")));
    var raced = new AtomicBoolean();
    var events = new ArrayList<ExecutionProgressEvent>();
    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor(
                    NativeInvocationContext.enterOwned(
                        dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
                .execute(
                    request,
                    bindings,
                    event -> {
                      events.add(event);
                      if (event.category() == ExecutionProgressEvent.Category.PERSIST
                          && event.status() == ExecutionProgressEvent.Status.STARTED
                          && !raced.getAndSet(true)) {
                        try {
                          Files.writeString(
                              output, "competing output", StandardOpenOption.CREATE_NEW);
                        } catch (IOException exception) {
                          throw new UncheckedIOException(exception);
                        }
                      }
                    }));
    assertTrue(
        raced.get(),
        () -> "Execution failed before the publication injection: " + failure.problem());
    assertEquals("competing output", Files.readString(output));
    assertEquals(4, failure.evidence().admission().operations().size());
    assertEquals(
        WorkbookExecutionEvidence.Backend.FULL_XSSF,
        failure.evidence().admission().executionStrategy().orElseThrow().backend());
    var input = failure.evidence().admission().materializedInputs().getFirst();
    assertEquals(root.resolve("label.txt").toString(), input.path());
    assertEquals(5, input.byteSize());
    assertInstanceOf(
        WorkbookExecutionEvidence.Structural.Established.class, failure.evidence().structural());
    assertEquals(
        List.of("check"),
        failure.assertions().stream().map(value -> value.stepId().value()).toList());
    assertEquals(1, failure.inspections().size());
    assertEquals(List.of(ExecutionProgressEvent.Status.FAILED), terminalPlanStatuses(events));
    assertEmptyScratch(root.resolve("scratch"));
  }

  @Test
  void doctorUnwindsPriorFileSnapshotWhenLaterLazyInputThrowsAnError() throws Exception {
    Files.writeString(root.resolve("label.txt"), "Label");
    WorkbookPlan request =
        plan(
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(
                ensureSheet(),
                setText("file", "A1", new TextSourceSpec.Utf8File("label.txt")),
                setText("stdin", "A2", new TextSourceSpec.Stdin())));
    ExecutionGrant base =
        ExecutionGrantTestSupport.forPlan(request, root, root.resolve("label.txt"));
    var resources = new ArrayList<>(base.readableResources());
    resources.add(new ExecutionGrant.ReadAuthority.Stdin());
    var grant =
        new ExecutionGrant(
            1024 * 1024,
            resources,
            base.operationIds(),
            base.targetAuthority(),
            base.publicationAuthority(),
            base.allowedSecretReferences(),
            base.acceptancePolicy());
    var marker = new AssertionError("synthetic opener failure");
    Path scratch = root.resolve("scratch");
    var snapshotObserved = new AtomicBoolean();
    var bindings =
        new ExecutionInputBindings(
            root,
            scratch,
            grant,
            Optional.of(() -> {
              try (var paths = Files.walk(scratch)) {
                assertTrue(
                    paths.anyMatch(Files::isRegularFile),
                    "A prior authorized input snapshot must exist before the injected Error");
              }
              snapshotObserved.set(true);
              throw marker;
            }),
            Optional.empty());
    assertSame(
        marker,
        assertThrows(
            AssertionError.class,
            () ->
                new GridGrindRequestDoctor(
                        NativeInvocationContext.enterOwned(
                            dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
                    .diagnose(request, bindings)));
    assertTrue(snapshotObserved.get());
    assertEquals("Label", Files.readString(root.resolve("label.txt")));
    assertEmptyScratch(scratch);
  }

  @Test
  void cleanupFailureReplacesProvisionalSuccessBeforeOneTerminalPlanEvent() {
    WorkbookPlan request = plan(new WorkbookPlan.WorkbookPersistence.None(), List.of());
    var events = new ArrayList<ExecutionProgressEvent>();
    var journal = ExecutionJournalRecorder.start(request, events::add, root);
    journal.buildSuccess(0);
    var provisional = WorkbookResults.success(List.of(), List.of(), List.of());
    var problem =
        GridGrindProblems.fromException(
            new IOException("synthetic cleanup failure"),
            new ProblemContext.ExecuteRequest(ExecutionRequestPaths.requestShape(request)));
    WorkbookResult failure =
        ExecutionResponseCloseFinalizer.withCloseFailure(
            provisional, request, journal, null, null, null, problem);
    assertTrue(terminalPlanStatuses(events).isEmpty());
    journal.completePlan(failure, 0);
    assertEquals(List.of(ExecutionProgressEvent.Status.FAILED), terminalPlanStatuses(events));
  }

  @Test
  void hostAssertionFailureRetainsEarlierPassedHostAssertionsInOrder() {
    WorkbookPlan request =
        plan(new WorkbookPlan.WorkbookPersistence.None(), List.of(ensureSheet()));
    ExecutionGrant base = ExecutionGrantTestSupport.forPlan(request, root);
    var policy =
        new AcceptancePolicy.RequireAll(
            List.of(
                new AcceptancePolicy.Requirement.TerminalAssertions(
                    List.of(
                        sheetPresent("host-pass", "Budget"),
                        sheetPresent("host-fail", "Missing")))));
    var grant =
        new ExecutionGrant(
            0,
            base.readableResources(),
            base.operationIds(),
            base.targetAuthority(),
            base.publicationAuthority(),
            base.allowedSecretReferences(),
            policy);
    var failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor(
                    NativeInvocationContext.enterOwned(
                        dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
                .execute(
                    request,
                    new ExecutionInputBindings(root, root.resolve("scratch"), grant),
                    ExecutionProgressSink.NOOP));
    var evidence =
        assertInstanceOf(
            WorkbookExecutionEvidence.TaskSpecific.Failed.class, failure.evidence().taskSpecific());
    assertEquals(
        List.of("host-pass", "host-fail"),
        evidence.assertions().stream().map(value -> value.result().stepId().value()).toList());
    assertTrue(
        evidence.assertions().stream()
            .allMatch(value ->
                value.origin() == WorkbookExecutionEvidence.TaskSpecific.Origin.HOST_REQUIRED));
    assertInstanceOf(AssertionResult.Passed.class, evidence.assertions().getFirst().result());
    assertInstanceOf(AssertionResult.Failed.class, evidence.assertions().getLast().result());
    assertTrue(failure.assertions().isEmpty());
  }

  private static WorkbookPlan plan(
      WorkbookPlan.WorkbookPersistence persistence, List<WorkbookStep> steps) {
    ExecutionPolicySpec defaults = ExecutionPolicySpec.defaults();
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        persistence,
        new ExecutionPolicySpec(
            defaults.mode(),
            new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE),
            defaults.calculation(),
            defaults.assertionMode()),
        FormulaEnvironmentSpec.empty(),
        steps);
  }

  private static MutationStep ensureSheet() {
    return new MutationStep(
        new StepId("ensure"),
        new SheetSelector.ByName(new SheetName("Budget")),
        new WorkbookMutationAction.EnsureSheet());
  }

  private static MutationStep setText(String id, String address, TextSourceSpec source) {
    return new MutationStep(
        new StepId(id),
        new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress(address)),
        new CellMutationAction.SetCell(new CellValueSpec.Text(source)));
  }

  private static AssertionStep sheetPresent(String id, String name) {
    return new AssertionStep(
        new StepId(id),
        new SheetSelector.ByName(new SheetName(name)),
        new PresenceAssertion.SheetPresent());
  }

  private static List<ExecutionProgressEvent.Status> terminalPlanStatuses(
      List<ExecutionProgressEvent> events) {
    return events.stream()
        .filter(event ->
            event.category() == ExecutionProgressEvent.Category.PLAN
                && event.status() != ExecutionProgressEvent.Status.STARTED)
        .map(ExecutionProgressEvent::status)
        .toList();
  }

  private static void assertEmptyScratch(Path scratch) throws IOException {
    if (Files.exists(scratch)) {
      try (var children = Files.list(scratch)) {
        assertFalse(
            children.findAny().isPresent(), "Invocation-owned scratch survived finalization");
      }
    }
  }
}
