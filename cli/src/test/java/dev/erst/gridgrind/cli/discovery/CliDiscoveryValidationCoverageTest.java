package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for CLI discovery validation helpers and task-definition accessors. */
class CliDiscoveryValidationCoverageTest {
  @Test
  void validatesPrimitiveCopyHelpersAndNullGuards() {
    GridGrindProtocolVersion protocolVersion = GridGrindProtocolVersion.current();
    TypeEntry typeEntry = GridGrindProtocolCatalog.catalog().plainTypes().getFirst().type();
    ShippedExampleEntry exampleEntry = GridGrindShippedExamples.catalog().examples().getFirst();
    TaskEntry taskEntry = GridGrindTaskCatalog.catalog().tasks().getFirst();
    TaskPhase taskPhase = taskEntry.workflow().phases().getFirst();
    TaskCapabilityRef capabilityRef = taskPhase.capabilityRefs().getFirst();

    assertEquals(protocolVersion, CliDiscoveryValidation.requireProtocolVersion(protocolVersion));
    assertEquals(
        List.of("alpha", "beta"),
        CliDiscoveryValidation.copyStrings(List.of("alpha", "beta"), "values"));
    assertEquals(
        List.of(typeEntry), CliDiscoveryValidation.copyTypeEntries(List.of(typeEntry), "entries"));
    assertEquals(
        List.of(taskEntry), CliDiscoveryValidation.copyTaskEntries(List.of(taskEntry), "tasks"));
    assertEquals(
        List.of(taskPhase), CliDiscoveryValidation.copyTaskPhases(List.of(taskPhase), "phases"));
    assertEquals(
        List.of(capabilityRef),
        CliDiscoveryValidation.copyTaskCapabilityRefs(List.of(capabilityRef), "capabilityRefs"));
    assertEquals(
        List.of(exampleEntry),
        CliDiscoveryValidation.copyExampleEntries(List.of(exampleEntry), "examples"));

    assertEquals(
        "protocolVersion must not be null",
        assertThrows(
                NullPointerException.class,
                () -> CliDiscoveryValidation.requireProtocolVersion(null))
            .getMessage());
    assertEquals(
        "entries must not be null",
        assertThrows(
                NullPointerException.class,
                () -> CliDiscoveryValidation.copyTypeEntries(null, "entries"))
            .getMessage());
    assertEquals(
        "entries must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    CliDiscoveryValidation.copyTypeEntries(
                        java.util.Arrays.asList((TypeEntry) null), "entries"))
            .getMessage());
  }

  @Test
  void taskDefinitionsExposeCanonicalEntriesAndRejectInvalidLookupIds() {
    TaskCatalog catalog = GridGrindTaskDefinitions.catalog();
    List<TaskEntry> entries = GridGrindTaskDefinitions.entries();

    assertFalse(entries.isEmpty());
    assertNotSame(entries, catalog.tasks());
    assertEquals(entries, catalog.tasks());
    assertTrue(GridGrindTaskDefinitions.entryFor("DASHBOARD").isPresent());
    assertTrue(GridGrindTaskDefinitions.entryFor("WORKBOOK_MAINTENANCE").isPresent());
    assertTrue(GridGrindTaskDefinitions.entryFor("NO_SUCH_TASK").isEmpty());
    GridGrindTaskDefinitions.validateCapabilityReferences();

    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindTaskDefinitions.entryFor(null))
            .getMessage());
    assertEquals(
        "id must not be blank",
        assertThrows(IllegalArgumentException.class, () -> GridGrindTaskDefinitions.entryFor("   "))
            .getMessage());
  }

  @Test
  void taskDefinitionsRejectDanglingCapabilityReferences() {
    TaskEntry brokenTask =
        new TaskEntry(
            "BROKEN",
            TaskTestFixtures.discoveryProfile("broken"),
            TaskTestFixtures.narrative("Broken task for direct validation coverage"),
            profile(),
            TaskTestFixtures.interactionProfile(),
            new TaskWorkflow(
                List.of(
                    new TaskPhase(
                        TaskPhasePurpose.AUTHOR,
                        "broken phase",
                        "exercise dangling capability validation",
                        List.of(new TaskCapabilityRef("mutationActionTypes", "NO_SUCH_ACTION")),
                        List.of())),
                List.of()));

    assertEquals(
        "Task BROKEN references unknown protocol capability mutationActionTypes:NO_SUCH_ACTION",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindTaskDefinitions.validateTaskCapabilityReferences(List.of(brokenTask)))
            .getMessage());
  }

  @Test
  void discoveryRecordsRejectInvalidMandatoryShapesAndNormalizeOptionalFailureFields() {
    assertEquals(
        "phases must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new TaskWorkflow(List.of(), List.of("pitfall")))
            .getMessage());

    assertEquals(
        "discoveryTerms must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskEntry(
                        "NO_TERMS",
                        new TaskDiscoveryProfile(
                            List.of(),
                            List.of("office"),
                            new TaskIntentProfile(
                                List.of(TaskGoalKind.AUTHOR), List.of(TaskArtifactKind.WORKBOOK))),
                        TaskTestFixtures.narrative("summary"),
                        profile(),
                        TaskTestFixtures.interactionProfile(),
                        TaskTestFixtures.workflow(
                            List.of(
                                new TaskPhase(
                                    TaskPhasePurpose.AUTHOR,
                                    "Phase",
                                    "Objective",
                                    List.of(new TaskCapabilityRef("sourceTypes", "NEW")),
                                    List.of("note"))))))
            .getMessage());

    assertEquals(
        "discoveryTerms must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskDiscoveryProfile(
                        List.of(),
                        List.of("office"),
                        new TaskIntentProfile(
                            List.of(TaskGoalKind.AUTHOR), List.of(TaskArtifactKind.WORKBOOK))))
            .getMessage());

    assertEquals(
        "goals must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new TaskIntentProfile(List.of(), List.of(TaskArtifactKind.WORKBOOK)))
            .getMessage());

    assertEquals(
        "artifacts must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new TaskIntentProfile(List.of(TaskGoalKind.AUTHOR), List.of()))
            .getMessage());

    assertEquals(
        "argument must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new CliFailureReport(
                        GridGrindProtocolVersion.current(),
                        2,
                        "print-task-plan",
                        GridGrindProblemCode.INVALID_ARGUMENTS,
                        "message",
                        CliFailureLocation.unavailable(),
                        null,
                        List.of(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "suggestions must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new CliFailureReport(
                        GridGrindProtocolVersion.current(),
                        2,
                        "print-task-plan",
                        GridGrindProblemCode.INVALID_ARGUMENTS,
                        "message",
                        CliFailureLocation.unavailable(),
                        Optional.empty(),
                        null,
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "resolution must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new CliFailureReport(
                        GridGrindProtocolVersion.current(),
                        2,
                        "print-task-plan",
                        GridGrindProblemCode.INVALID_ARGUMENTS,
                        "message",
                        CliFailureLocation.unavailable(),
                        Optional.empty(),
                        List.of(),
                        null))
            .getMessage());
    assertEquals(
        "exitCode must be positive",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CliFailureReport(
                        GridGrindProtocolVersion.current(),
                        0,
                        "print-task-plan",
                        GridGrindProblemCode.INVALID_ARGUMENTS,
                        "message",
                        CliFailureLocation.unavailable(),
                        Optional.empty(),
                        List.of(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "jsonLine and jsonColumn must either both be present or both be absent",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CliFailureLocation(
                        Optional.of("steps[0]"), Optional.of(1), Optional.empty()))
            .getMessage());
    assertEquals(
        "jsonColumn must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CliFailureLocation(Optional.of("steps[0]"), Optional.of(1), Optional.of(0)))
            .getMessage());
    assertEquals(
        "suggestedRequestPath must start with examples/",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ShippedExampleEntry(
                        "BROKEN",
                        Path.of("budget-request.json").toString(),
                        "summary",
                        ExampleWorkspaceMode.SELF_CONTAINED,
                        List.of()))
            .getMessage());
  }

  private static TaskExecutionProfile profile() {
    return TaskTestFixtures.profile();
  }
}
