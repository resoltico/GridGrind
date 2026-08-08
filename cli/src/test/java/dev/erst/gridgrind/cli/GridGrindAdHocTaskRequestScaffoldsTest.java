package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskCapabilityRef;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.discovery.TaskExecutionProfile;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhase;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.cli.discovery.TaskTestFixtures;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.JsonNodeFactory;
import tools.jackson.databind.node.ObjectNode;

/** Tests for generic ad hoc request scaffolds built from unpublished task descriptors. */
class GridGrindAdHocTaskRequestScaffoldsTest {
  @Test
  void scaffolderRejectsPublishedRecipeIds() {
    TaskEntry publishedTask = GridGrindCliRecipeRegistry.taskEntryFor("DASHBOARD").orElseThrow();
    TaskEntry exampleNamedTask =
        task(
            "WORKBOOK_HEALTH",
            new TaskExecutionProfile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.NONE,
                TaskMutationMode.READ_ONLY,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("inspectionQueryTypes", "GET_CELLS")))));

    IllegalArgumentException publishedTaskFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindAdHocTaskRequestScaffolds.requestFor(publishedTask));
    assertEquals(
        "Published recipe id DASHBOARD must use the canonical recipe registry",
        publishedTaskFailure.getMessage());

    IllegalArgumentException publishedExampleFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindAdHocTaskRequestScaffolds.requestFor(exampleNamedTask));
    assertEquals(
        "Published recipe id WORKBOOK_HEALTH must use the canonical recipe registry",
        publishedExampleFailure.getMessage());
  }

  @Test
  void scaffolderRejectsIncompatibleOverwriteDefaults() {
    TaskEntry invalidOverwriteTask =
        task(
            "INVALID_OVERWRITE",
            new TaskExecutionProfile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.OVERWRITE,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")))));
    IllegalStateException invalidOverwrite =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindAdHocTaskRequestScaffolds.requestFor(invalidOverwriteTask));
    assertTrue(invalidOverwrite.getMessage().contains("cannot plan OVERWRITE persistence"));
  }

  @Test
  void scaffolderBuildsOverwriteStarterRequestForExistingWorkbookTasks() {
    TaskEntry overwriteTask =
        task(
            "OVERWRITE_EXISTING",
            new TaskExecutionProfile(
                TaskSourceMode.EXISTING_WORKBOOK,
                TaskPersistenceMode.OVERWRITE,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")))));

    WorkbookPlan request = GridGrindAdHocTaskRequestScaffolds.requestFor(overwriteTask);

    assertEquals("EXISTING", sourceType(request));
    assertTrue(inputPath(request).endsWith(".xlsx"));
    assertTrue(inputPath(request).contains("overwrite-existing"));
    assertEquals("OVERWRITE", persistenceType(request));
    assertFalse(steps(request).isEmpty());
  }

  @Test
  void scaffolderBuildsGenericRequestsForAdHocNoneAndSaveAsTasks() {
    TaskEntry noneTask =
        task(
            "AD_HOC_DISCOVERY",
            new TaskExecutionProfile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.NONE,
                TaskMutationMode.READ_ONLY,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("inspectionQueryTypes", "GET_CELLS")))));
    TaskEntry saveAsTask =
        task(
            "AD_HOC_EXPORT",
            new TaskExecutionProfile(
                TaskSourceMode.EXISTING_WORKBOOK,
                TaskPersistenceMode.SAVE_AS,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")))));

    WorkbookPlan noneRequest = GridGrindAdHocTaskRequestScaffolds.requestFor(noneTask);
    WorkbookPlan saveAsRequest = GridGrindAdHocTaskRequestScaffolds.requestFor(saveAsTask);

    assertEquals("NEW", sourceType(noneRequest));
    assertEquals("NONE", persistenceType(noneRequest));
    assertFalse(steps(noneRequest).isEmpty());

    assertEquals("EXISTING", sourceType(saveAsRequest));
    assertEquals("starter-ad-hoc-export-input.xlsx", inputPath(saveAsRequest));
    assertEquals("SAVE_AS", persistenceType(saveAsRequest));
    assertEquals("REPLACE", persistenceIfExists(saveAsRequest));
    assertEquals("starter-ad-hoc-export-output.xlsx", outputPath(saveAsRequest));
    assertFalse(steps(saveAsRequest).isEmpty());
  }

  @Test
  void scaffolderDeduplicatesRepeatedCapabilityTemplatesWithinOnePhase() {
    TaskEntry duplicateCapabilityTask =
        task(
            "AD_HOC_DUPLICATE_CAPABILITY",
            new TaskExecutionProfile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.NONE,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(
                new TaskPhase(
                    dev.erst.gridgrind.cli.discovery.TaskPhasePurpose.AUTHOR,
                    "Phase One",
                    "First objective",
                    List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")),
                    List.of("note one")),
                new TaskPhase(
                    dev.erst.gridgrind.cli.discovery.TaskPhasePurpose.AUTHOR,
                    "Phase Two",
                    "Second objective",
                    List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")),
                    List.of("note two"))));

    WorkbookPlan request = GridGrindAdHocTaskRequestScaffolds.requestFor(duplicateCapabilityTask);

    assertEquals(1, steps(request).size());
    assertEquals("phase-1-step-1", textField(firstStep(request), "stepId"));
  }

  @Test
  void scaffolderSkipsCapabilitiesWithoutPublishedStepTemplates() {
    TaskEntry partiallyPublishedTask =
        task(
            "AD_HOC_PARTIAL_DISCOVERY",
            new TaskExecutionProfile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.NONE,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(
                phase(
                    List.of(
                        new TaskCapabilityRef("mutationActionTypes", "SET_CELL"),
                        new TaskCapabilityRef("mutationActionTypes", "NO_SUCH_CAPABILITY")))));

    WorkbookPlan request = GridGrindAdHocTaskRequestScaffolds.requestFor(partiallyPublishedTask);

    assertEquals(1, steps(request).size());
    assertEquals("phase-1-step-1", textField(firstStep(request), "stepId"));
  }

  @Test
  void scaffolderRejectsNullTasksAndCarriesExternalAssetRequests() {
    NullPointerException nullTask =
        assertThrows(
            NullPointerException.class,
            () -> GridGrindAdHocTaskRequestScaffolds.requestFor((TaskEntry) null));
    assertEquals("task must not be null", nullTask.getMessage());

    WorkbookPlan externalPayloadRequest =
        GridGrindAdHocTaskRequestScaffolds.requestFor(
            task(
                "EXTERNAL_PAYLOAD_IMPORT",
                new TaskExecutionProfile(
                    TaskSourceMode.EXISTING_WORKBOOK,
                    TaskPersistenceMode.SAVE_AS,
                    TaskMutationMode.MUTATING,
                    TaskAssetMode.REQUIRES_EXTERNAL_PAYLOADS),
                List.of(
                    phase(
                        List.of(
                            new TaskCapabilityRef(
                                "mutationActionTypes", "IMPORT_CUSTOM_XML_MAPPING"))))));

    assertEquals("EXISTING", sourceType(externalPayloadRequest));
    assertEquals("SAVE_AS", persistenceType(externalPayloadRequest));
    assertEquals("REPLACE", persistenceIfExists(externalPayloadRequest));
  }

  @Test
  void scaffolderWrapsSerializationFailuresFromInvalidGeneratedTrees() {
    ObjectNode cyclic = JsonNodeFactory.instance.objectNode();
    cyclic.put("protocolVersion", "V2");
    cyclic.set("self", cyclic);

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindAdHocTaskRequestScaffolds.decodedRequest("BROKEN_TASK", cyclic));

    assertEquals("CLI-generated task request is invalid for BROKEN_TASK", failure.getMessage());
  }

  @Test
  void decodedRequestRejectsBlankTaskIdsWhenGeneratedJsonIsInvalid() {
    ObjectNode cyclic = JsonNodeFactory.instance.objectNode();
    cyclic.put("protocolVersion", "V2");
    cyclic.set("self", cyclic);

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindAdHocTaskRequestScaffolds.decodedRequest(" ", cyclic));

    assertEquals("taskId must not be blank", failure.getMessage());
  }

  private static TaskEntry task(
      String id, TaskExecutionProfile executionProfile, List<TaskPhase> phases) {
    return TaskTestFixtures.task(id, executionProfile, phases);
  }

  private static TaskPhase phase(List<TaskCapabilityRef> capabilityRefs) {
    return TaskTestFixtures.phase(capabilityRefs);
  }

  private static String sourceType(WorkbookPlan request) {
    return textField(requestTree(request).path("source"), "type");
  }

  private static String persistenceType(WorkbookPlan request) {
    return textField(requestTree(request).path("persistence"), "type");
  }

  private static String inputPath(WorkbookPlan request) {
    return textField(requestTree(request).path("source"), "path");
  }

  private static String persistenceIfExists(WorkbookPlan request) {
    return textField(requestTree(request).path("persistence"), "ifExists");
  }

  private static String outputPath(WorkbookPlan request) {
    return textField(requestTree(request).path("persistence"), "path");
  }

  private static List<JsonNode> steps(WorkbookPlan request) {
    List<JsonNode> steps = new java.util.ArrayList<>();
    for (JsonNode step : requestTree(request).path("steps")) {
      steps.add(step);
    }
    return List.copyOf(steps);
  }

  private static JsonNode firstStep(WorkbookPlan request) {
    return steps(request).getFirst();
  }

  private static ObjectNode requestTree(WorkbookPlan request) {
    return GridGrindJsonOutput.requestTree(request);
  }

  private static String textField(JsonNode node, String fieldName) {
    return java.util.Objects.requireNonNull(node.path(fieldName).stringValue());
  }
}
