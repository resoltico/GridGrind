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

/** Task descriptor for workbook maintenance and sheet-copy workflows. */
final class WorkbookMaintenanceTaskDefinition {
  private WorkbookMaintenanceTaskDefinition() {}

  static TaskEntry entry() {
    return task(
        "WORKBOOK_MAINTENANCE",
        discovery(
            List.of("maintenance", "copy sheet", "repair", "normalize workbook", "sheet clone"),
            List.of("office", "maintenance", "repair", "copy", "comments"),
            intent(
                List.of(TaskGoalKind.INSPECT, TaskGoalKind.MAINTAIN, TaskGoalKind.VERIFY),
                List.of(
                    TaskArtifactKind.WORKBOOK,
                    TaskArtifactKind.SHEET,
                    TaskArtifactKind.COMMENT,
                    TaskArtifactKind.DRAWING_OBJECT))),
        narrative(
            "Safely copy or normalize existing workbook sheets while verifying comments,"
                + " drawings, and other cloned workbook state after the change.",
            List.of(
                "Workbook maintenance flows can inspect the source before making structural"
                    + " changes.",
                "Copy-sheet work can be followed immediately by comment and drawing readback.",
                "The updated workbook can be saved separately while you verify the result."),
            List.of(
                "Existing workbook path.",
                "Source sheet names and destination copy names.",
                "Persistence target when the maintained workbook should be kept."),
            List.of(
                "Workbook-level findings after the maintenance pass.",
                "Comment rereads on copied sheets.",
                "Drawing-object rereads on copied sheets when they matter.")),
        profile(
            TaskSourceMode.EXISTING_WORKBOOK,
            TaskPersistenceMode.SAVE_AS,
            TaskMutationMode.MUTATING,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(
                TaskInputKind.SOURCE_WORKBOOK_PATH,
                TaskInputKind.TARGET_SHEET_NAMES,
                TaskInputKind.PERSISTENCE_TARGET_PATH),
            List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.HEALTH_ANALYSIS)),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.INSPECT,
                    "Inspect The Current Workbook",
                    "Read source workbook facts before changing sheet structure.",
                    List.of(
                        ref("sourceTypes", "EXISTING"),
                        ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY"),
                        ref("inspectionQueryTypes", "GET_COMMENTS")),
                    List.of("Factual inspection is the baseline for any maintenance pass.")),
                phase(
                    TaskPhasePurpose.AUTHOR,
                    "Apply Structural Maintenance",
                    "Copy one existing sheet into a new workbook-visible destination sheet.",
                    List.of(
                        ref("mutationActionTypes", "COPY_SHEET"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Copy-sheet work is safest when the destination name is explicit.")),
                phase(
                    TaskPhasePurpose.VERIFY,
                    "Verify The Copied Surface",
                    "Read back copied comments, drawings, and workbook findings immediately.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_COMMENTS"),
                        ref("inspectionQueryTypes", "GET_DRAWING_OBJECTS"),
                        ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS")),
                    List.of("Verification closes the loop on maintenance changes."))),
            List.of(
                "Replace placeholder sheet names before execution; copy-sheet requires an"
                    + " existing source sheet.",
                "Maintenance flows are safest with SAVE_AS so the source workbook stays intact"
                    + " until verification passes.",
                "Use targeted rereads after COPY_SHEET instead of assuming copied comments or"
                    + " drawings stayed coherent.")));
  }
}
