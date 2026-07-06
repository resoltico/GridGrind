package dev.erst.gridgrind.cli.examples;

import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.assetBackedStarter;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.discovery;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.intent;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.narrative;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.phase;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.profile;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.ref;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.signals;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.task;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.workflow;

import dev.erst.gridgrind.cli.discovery.TaskArtifactKind;
import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskGoalKind;
import dev.erst.gridgrind.cli.discovery.TaskInputKind;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.cli.discovery.TaskVerificationKind;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import java.util.List;

/** Published task recipe for workbook maintenance and sheet-copy workflows. */
final class WorkbookMaintenanceTaskRecipe {
  private WorkbookMaintenanceTaskRecipe() {}

  static GridGrindTaskRecipeDefinition definition() {
    return GridGrindTaskRecipeSupport.definition(
        task(
            "WORKBOOK_MAINTENANCE",
            assetBackedStarter(
                "WORKBOOK_MAINTENANCE", "task-starter-assets/workbook-ops-source.xlsx"),
            List.of("office", "maintenance", "repair", "copy", "comments"),
            discovery(
                List.of("maintenance", "copy sheet", "repair", "normalize workbook", "sheet clone"),
                intent(
                    List.of(TaskGoalKind.INSPECT, TaskGoalKind.MAINTAIN, TaskGoalKind.VERIFY),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.SHEET,
                        TaskArtifactKind.COMMENT,
                        TaskArtifactKind.DRAWING_OBJECT))),
            narrative(
                "Safely copy or normalize existing workbook sheets while verifying comments, drawings, and other cloned workbook state after the change.",
                List.of(
                    "Workbook maintenance flows can inspect the source before making structural changes.",
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
                        List.of(
                            "Copy-sheet work is safest when the destination name is explicit.")),
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
                    "Replace placeholder sheet names before execution; copy-sheet requires an existing source sheet.",
                    "Maintenance flows are safest with SAVE_AS so the source workbook stays intact until verification passes.",
                    "Use targeted rereads after COPY_SHEET instead of assuming copied comments or drawings stayed coherent."))),
        starterPlan());
  }

  private static WorkbookPlan starterPlan() {
    String taskId = "WORKBOOK_MAINTENANCE";
    String sourceAsset = TaskStarterRecipeSupport.taskStarterAsset("workbook-ops-source.xlsx");
    return ExampleWorkbookPlans.defaultExecutionPlan(
        TaskStarterRecipeSupport.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.ExistingFile(sourceAsset),
        ExampleWorkbookPlans.saveAs(TaskStarterRecipeSupport.taskWorkbookPath(taskId)),
        ExampleSteps.read(
            "read-source-workbook",
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary()),
        ExampleSteps.step(
            "copy-template-sheet",
            ExampleSelectors.sheet("Template"),
            new WorkbookMutationAction.CopySheet("Template Copy")),
        ExampleSteps.read(
            "read-copied-comments",
            new dev.erst.gridgrind.contract.selector.CellSelector.AllUsedInSheet("Template Copy"),
            new SheetIntrospectionQuery.GetComments()),
        ExampleSteps.read(
            "read-copied-drawing-objects",
            new DrawingObjectSelector.AllOnSheet("Template Copy"),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
        ExampleSteps.read(
            "read-maintenance-findings",
            ExampleSelectors.workbook(),
            new InspectionAnalysisQuery.AnalyzeWorkbookFindings()));
  }
}
