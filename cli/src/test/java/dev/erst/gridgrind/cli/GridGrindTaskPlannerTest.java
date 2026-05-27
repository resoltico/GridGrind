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
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.JsonNodeFactory;
import tools.jackson.databind.node.ObjectNode;

/** Tests for CLI-owned starter requests derived from task descriptors. */
class GridGrindTaskPlannerTest {
  @Test
  void plannerBuildsDashboardStarterRequest() {
    WorkbookPlan request = GridGrindTaskPlanner.requestFor("DASHBOARD");

    assertEquals("NEW", sourceType(request));
    assertEquals("SAVE_AS", persistenceType(request));
    assertTrue(outputPath(request).endsWith(".xlsx"));
    assertTrue(outputPath(request).contains("dashboard"));
    assertFalse(steps(request).isEmpty());
    assertTrue(firstStep(request).has("target"));
    assertTrue(firstStep(request).has("action") || firstStep(request).has("query"));
    assertEquals("ensure-ops", textField(firstStep(request), "stepId"));
    assertTrue(GridGrindEngine.requestDoctor().diagnose(request).valid());
  }

  @Test
  void plannerBuildsAuditStarterRequest() {
    WorkbookPlan request = GridGrindTaskPlanner.requestFor("AUDIT_EXISTING_WORKBOOK");

    assertEquals("EXISTING", sourceType(request));
    assertTrue(inputPath(request).endsWith(".xlsx"));
    assertTrue(inputPath(request).contains("task-starter-assets/workbook-ops-source.xlsx"));
    assertEquals("NONE", persistenceType(request));
    assertFalse(steps(request).isEmpty());
    assertTrue(firstStep(request).has("query"));
  }

  @Test
  void plannerCarriesExistingWorkbookAndExternalPayloadPlaceholders() {
    WorkbookPlan customXml = GridGrindTaskPlanner.requestFor("CUSTOM_XML_WORKFLOW");
    WorkbookPlan maintenance = GridGrindTaskPlanner.requestFor("WORKBOOK_MAINTENANCE");

    assertEquals("EXISTING", sourceType(customXml));
    assertEquals("SAVE_AS", persistenceType(customXml));
    assertFalse(steps(customXml).isEmpty());
    assertEquals("generated-workbooks/custom-xml-workflow.xlsx", outputPath(customXml));

    assertEquals("EXISTING", sourceType(maintenance));
    assertEquals("SAVE_AS", persistenceType(maintenance));
    assertFalse(steps(maintenance).isEmpty());
    assertTrue(outputPath(maintenance).contains("workbook-maintenance"));
    assertTrue(inputPath(maintenance).contains("task-starter-assets/workbook-ops-source.xlsx"));
  }

  @Test
  void plannerRejectsUnknownAndIncompatibleTaskDefaults() {
    IllegalArgumentException unknownTask =
        assertThrows(
            IllegalArgumentException.class, () -> GridGrindTaskPlanner.requestFor("BOGUS_TASK"));
    assertTrue(unknownTask.getMessage().contains("Unknown task id"));

    TaskEntry invalidOverwriteTask =
        task(
            "INVALID_OVERWRITE",
            new TaskExecutionProfile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.OVERWRITE_SOURCE,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")))));
    IllegalStateException invalidOverwrite =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindTaskPlanner.requestFor(invalidOverwriteTask));
    assertTrue(invalidOverwrite.getMessage().contains("cannot plan OVERWRITE persistence"));
  }

  @Test
  void plannerRejectsNullAndBlankTaskIds() {
    NullPointerException nullTaskId =
        assertThrows(
            NullPointerException.class, () -> GridGrindTaskPlanner.requestFor((String) null));
    assertEquals("taskId must not be null", nullTaskId.getMessage());

    IllegalArgumentException blankTaskId =
        assertThrows(IllegalArgumentException.class, () -> GridGrindTaskPlanner.requestFor(" "));
    assertEquals("taskId must not be blank", blankTaskId.getMessage());
  }

  @Test
  void plannerBuildsOverwriteStarterRequestForExistingWorkbookTasks() {
    TaskEntry overwriteTask =
        task(
            "OVERWRITE_EXISTING",
            new TaskExecutionProfile(
                TaskSourceMode.EXISTING_WORKBOOK,
                TaskPersistenceMode.OVERWRITE_SOURCE,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            List.of(phase(List.of(new TaskCapabilityRef("mutationActionTypes", "SET_CELL")))));

    WorkbookPlan request = GridGrindTaskPlanner.requestFor(overwriteTask);

    assertEquals("EXISTING", sourceType(request));
    assertTrue(inputPath(request).endsWith(".xlsx"));
    assertTrue(inputPath(request).contains("overwrite-existing"));
    assertEquals("OVERWRITE", persistenceType(request));
    assertFalse(steps(request).isEmpty());
  }

  @Test
  void plannerBuildsGenericRequestsForAdHocNoneAndSaveAsTasks() {
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

    WorkbookPlan noneRequest = GridGrindTaskPlanner.requestFor(noneTask);
    WorkbookPlan saveAsRequest = GridGrindTaskPlanner.requestFor(saveAsTask);

    assertEquals("NEW", sourceType(noneRequest));
    assertEquals("NONE", persistenceType(noneRequest));
    assertFalse(steps(noneRequest).isEmpty());

    assertEquals("EXISTING", sourceType(saveAsRequest));
    assertEquals("starter-ad-hoc-export-input.xlsx", inputPath(saveAsRequest));
    assertEquals("SAVE_AS", persistenceType(saveAsRequest));
    assertEquals("starter-ad-hoc-export-output.xlsx", outputPath(saveAsRequest));
    assertFalse(steps(saveAsRequest).isEmpty());
  }

  @Test
  void plannerDeduplicatesRepeatedCapabilityTemplatesWithinOnePhase() {
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

    WorkbookPlan request = GridGrindTaskPlanner.requestFor(duplicateCapabilityTask);

    assertEquals(1, steps(request).size());
    assertEquals("phase-1-step-1", textField(firstStep(request), "stepId"));
  }

  @Test
  void plannerSkipsCapabilitiesWithoutPublishedStepTemplates() {
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

    WorkbookPlan request = GridGrindTaskPlanner.requestFor(partiallyPublishedTask);

    assertEquals(1, steps(request).size());
    assertEquals("phase-1-step-1", textField(firstStep(request), "stepId"));
  }

  @Test
  void plannerRejectsNullTasksAndCarriesExternalAssetRequests() {
    NullPointerException nullTask =
        assertThrows(
            NullPointerException.class, () -> GridGrindTaskPlanner.requestFor((TaskEntry) null));
    assertEquals("task must not be null", nullTask.getMessage());

    WorkbookPlan externalPayloadRequest =
        GridGrindTaskPlanner.requestFor(
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
  }

  @Test
  void plannerWrapsSerializationFailuresFromInvalidGeneratedTrees() {
    ObjectNode cyclic = JsonNodeFactory.instance.objectNode();
    cyclic.put("protocolVersion", "V1");
    cyclic.set("self", cyclic);

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindTaskPlanner.decodedRequest("BROKEN_TASK", cyclic));

    assertEquals("CLI-generated task request is invalid for BROKEN_TASK", failure.getMessage());
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
    return GridGrindJson.requestTree(request);
  }

  private static String textField(JsonNode node, String fieldName) {
    return java.util.Objects.requireNonNull(node.path(fieldName).stringValue());
  }
}
