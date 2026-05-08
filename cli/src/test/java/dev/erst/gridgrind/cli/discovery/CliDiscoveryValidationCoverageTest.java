package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for CLI discovery validation helpers and task-definition accessors. */
class CliDiscoveryValidationCoverageTest {
  @Test
  void validatesPrimitiveCopyHelpersAndNullGuards() {
    GridGrindProtocolVersion protocolVersion = GridGrindProtocolVersion.current();
    WorkbookPlan requestTemplate =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    TypeEntry typeEntry = GridGrindProtocolCatalog.catalog().plainTypes().getFirst().type();
    ShippedExampleEntry exampleEntry = GridGrindShippedExamples.catalog().examples().getFirst();
    TaskEntry taskEntry = GridGrindTaskCatalog.catalog().tasks().getFirst();
    TaskPhase taskPhase = taskEntry.phases().getFirst();
    TaskCapabilityRef capabilityRef = taskPhase.capabilityRefs().getFirst();

    assertEquals(protocolVersion, CliDiscoveryValidation.requireProtocolVersion(protocolVersion));
    assertEquals(requestTemplate, CliDiscoveryValidation.requireRequestTemplate(requestTemplate));
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
        "requestTemplate must not be null",
        assertThrows(
                NullPointerException.class,
                () -> CliDiscoveryValidation.requireRequestTemplate(null))
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
            "Broken task for direct validation coverage",
            profile(),
            List.of(),
            List.of(),
            List.of("broken"),
            List.of("none"),
            List.of("input"),
            List.of(),
            List.of(
                new TaskPhase(
                    TaskPhasePurpose.AUTHOR,
                    "broken phase",
                    "exercise dangling capability validation",
                    List.of(new TaskCapabilityRef("mutationActionTypes", "NO_SUCH_ACTION")),
                    List.of())),
            List.of());

    assertEquals(
        "Task BROKEN references unknown protocol capability mutationActionTypes:NO_SUCH_ACTION",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindTaskDefinitions.validateTaskCapabilityReferences(List.of(brokenTask)))
            .getMessage());
  }

  private static TaskExecutionProfile profile() {
    return new TaskExecutionProfile(
        TaskSourceMode.NEW_WORKBOOK,
        TaskPersistenceMode.NONE,
        TaskMutationMode.MUTATING,
        TaskAssetMode.SELF_CONTAINED);
  }
}
