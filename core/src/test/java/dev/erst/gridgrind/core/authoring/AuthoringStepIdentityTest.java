package dev.erst.gridgrind.core.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.plan.PresenceAssertion;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Checks globally occupied step IDs and failed-append isolation in Java authoring. */
class AuthoringStepIdentityTest {
  private static PlannedMutation mutation() {
    return new PlannedMutation(
        new SheetSelector.ByName(new SheetName("Sheet")), new WorkbookMutationAction.EnsureSheet());
  }

  private static PlannedInspection inspection() {
    return new PlannedInspection(
        new WorkbookSelector.Current(), new WorkbookIntrospectionQuery.GetWorkbookSummary());
  }

  private static PlannedAssertion assertion() {
    return new PlannedAssertion(
        new SheetSelector.ByName(new SheetName("Sheet")), new PresenceAssertion.SheetPresent());
  }

  @Test
  void continuationAllocatesAgainstGlobalOccupiedIdsRatherThanFamilyCounts() {
    var original =
        GridGrindPlan.newWorkbook()
            .inspect(inspection().named(new StepId("mutation-001")))
            .assertThat(assertion().named(new StepId("inspection-001")))
            .mutate(mutation().named(new StepId("assertion-001")))
            .mutate(mutation().named(new StepId("mutation-003")))
            .toPlan();
    var continued =
        GridGrindPlan.from(original)
            .mutate(mutation())
            .inspect(inspection())
            .assertThat(assertion())
            .mutate(mutation())
            .toPlan();
    assertEquals(
        List.of(
            "mutation-001",
            "inspection-001",
            "assertion-001",
            "mutation-003",
            "mutation-002",
            "inspection-002",
            "assertion-002",
            "mutation-004"),
        continued.steps().stream().map(step -> step.stepId().value()).toList());
    assertEquals(4, original.steps().size());
  }

  @Test
  void failedExplicitAppendLeavesBuilderAndAllocationUnchanged() {
    var builder = GridGrindPlan.newWorkbook().mutate(mutation().named(new StepId("occupied")));
    assertThrows(
        InvalidRequestException.class,
        () -> builder.inspect(inspection().named(new StepId("occupied"))));
    assertThrows(
        InvalidRequestException.class,
        () -> builder.assertThat(assertion().named(new StepId("occupied"))));
    assertThrows(
        InvalidRequestException.class,
        () -> builder.mutate(mutation().named(new StepId("occupied"))));
    assertEquals(
        List.of("occupied", "mutation-001", "inspection-001", "assertion-001"),
        builder
            .mutate(mutation())
            .inspect(inspection())
            .assertThat(assertion())
            .toPlan()
            .steps()
            .stream()
            .map(step -> step.stepId().value())
            .toList());
  }

  @Test
  void generatedIdsSkipLaterCrossFamilyInsertionsAndKeepFailedAppendsStable() {
    var builder =
        GridGrindPlan.newWorkbook()
            .mutate(mutation())
            .inspect(inspection().named(new StepId("mutation-002")))
            .mutate(mutation());
    var beforeFailure = builder.toPlan();
    assertThrows(
        InvalidRequestException.class,
        () -> builder.mutate(mutation().named(new StepId("mutation-002"))));
    assertEquals(beforeFailure, builder.toPlan());
    assertEquals(
        List.of("mutation-001", "mutation-002", "mutation-003", "inspection-001", "mutation-004"),
        builder.inspect(inspection()).mutate(mutation()).toPlan().steps().stream()
            .map(step -> step.stepId().value())
            .toList());
  }
}
