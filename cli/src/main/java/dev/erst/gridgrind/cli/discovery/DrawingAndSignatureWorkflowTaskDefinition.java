package dev.erst.gridgrind.cli.discovery;

import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.discovery;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.intent;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.narrative;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.phase;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.profile;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.ref;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.signals;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.task;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.workflow;

import java.util.List;

/** Task descriptor for drawing and signature-line workflows. */
final class DrawingAndSignatureWorkflowTaskDefinition {
  private DrawingAndSignatureWorkflowTaskDefinition() {}

  static TaskEntry entry() {
    return task(
        "DRAWING_AND_SIGNATURE_WORKFLOW",
        discovery(
            List.of("drawing", "signature line", "picture", "shape", "embedded object"),
            List.of("office", "drawing", "signature", "picture", "embedded-object"),
            intent(
                List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY),
                List.of(
                    TaskArtifactKind.WORKBOOK,
                    TaskArtifactKind.DRAWING_OBJECT,
                    TaskArtifactKind.SIGNATURE_LINE))),
        narrative(
            "Author or inspect drawing-backed workbook content such as signature lines, pictures,"
                + " shapes, and embedded objects.",
            List.of(
                "Drawing-backed workbook objects are authored by name instead of fragile XML"
                    + " surgery.",
                "Anchors can be moved authoritatively after creation.",
                "Drawing-object readback confirms what the workbook now contains."),
            List.of(
                "Target sheet and object names.",
                "Anchor positions.",
                "Persistence target when the workbook should be saved."),
            List.of(
                "Signature-line preview images.",
                "Follow-up drawing-object payload extraction or deletion.",
                "Picture, shape, or embedded-object variants alongside signature lines.")),
        profile(
            TaskSourceMode.NEW_WORKBOOK,
            TaskPersistenceMode.SAVE_AS,
            TaskMutationMode.MUTATING,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(
                TaskInputKind.TARGET_SHEET_NAMES,
                TaskInputKind.TARGET_OBJECT_NAMES,
                TaskInputKind.DRAWING_ANCHORS,
                TaskInputKind.PERSISTENCE_TARGET_PATH),
            List.of(TaskVerificationKind.FACT_READBACK)),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.PREPARE,
                    "Prepare The Drawing Surface",
                    "Create the target sheet before any named drawing object is authored.",
                    List.of(ref("sourceTypes", "NEW"), ref("mutationActionTypes", "ENSURE_SHEET")),
                    List.of("A stable sheet name gives drawing objects a durable home.")),
                phase(
                    TaskPhasePurpose.AUTHOR,
                    "Author Drawing Objects",
                    "Create one named drawing object and move it authoritatively when needed.",
                    List.of(
                        ref("mutationActionTypes", "SET_SIGNATURE_LINE"),
                        ref("mutationActionTypes", "SET_DRAWING_OBJECT_ANCHOR"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Named drawing objects are easier to inspect and move later.")),
                phase(
                    TaskPhasePurpose.VERIFY,
                    "Inspect The Result",
                    "Read back worksheet drawing metadata after authoring or anchor changes.",
                    List.of(ref("inspectionQueryTypes", "GET_DRAWING_OBJECTS")),
                    List.of("Use factual reread to confirm the visible workbook surface."))),
            List.of(
                "Signature lines are VML-backed drawing objects with workbook-side constraints.",
                "Invalid drawing payloads are rejected without leaving partial artifacts behind.",
                "Image-heavy workflows should prefer file-backed binary sources over huge inline"
                    + " base64 payloads.")));
  }
}
