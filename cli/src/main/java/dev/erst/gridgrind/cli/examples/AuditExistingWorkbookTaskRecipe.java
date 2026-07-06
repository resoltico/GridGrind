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
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.InspectionSurfaceQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import java.util.List;

/** Published task recipe for read-only workbook audit workflows. */
final class AuditExistingWorkbookTaskRecipe {
  private AuditExistingWorkbookTaskRecipe() {}

  static GridGrindTaskRecipeDefinition definition() {
    return GridGrindTaskRecipeSupport.definition(
        task(
            "AUDIT_EXISTING_WORKBOOK",
            assetBackedStarter(
                "AUDIT_EXISTING_WORKBOOK", "task-starter-assets/workbook-ops-source.xlsx"),
            List.of("office", "audit", "analysis", "readback", "safety"),
            discovery(
                List.of("audit", "inspect", "health", "findings", "diagnose", "read only"),
                intent(
                    List.of(TaskGoalKind.INSPECT, TaskGoalKind.ANALYZE, TaskGoalKind.VERIFY),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.PACKAGE_SECURITY,
                        TaskArtifactKind.FORMULA_SURFACE))),
            narrative(
                "Inspect an existing workbook, surface health findings, and avoid mutation unless a follow-up workflow explicitly asks for it.",
                List.of(
                    "Existing workbook structure is surfaced as facts instead of guesses.",
                    "Health analyses aggregate into one workbook-level findings view.",
                    "The default audit posture stays read-only and non-persistent."),
                List.of(
                    "Source workbook path.",
                    "Target sheets or workbook areas that matter to the audit.",
                    "Any expected invariants that should later become assertions."),
                List.of(
                    "Package security inspection for protected OOXML files.",
                    "Targeted formula-surface inspection before broad workbook findings.",
                    "Follow-up assertion plans after the first audit pass.")),
            profile(
                TaskSourceMode.EXISTING_WORKBOOK,
                TaskPersistenceMode.NONE,
                TaskMutationMode.READ_ONLY,
                TaskAssetMode.SELF_CONTAINED),
            signals(
                List.of(TaskInputKind.SOURCE_WORKBOOK_PATH, TaskInputKind.TARGET_SHEET_NAMES),
                List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.HEALTH_ANALYSIS)),
            workflow(
                List.of(
                    phase(
                        TaskPhasePurpose.INSPECT,
                        "Open And Inspect The Package",
                        "Start with workbook-level facts and security state before diving into sheet logic.",
                        List.of(
                            ref("sourceTypes", "EXISTING"),
                            ref("persistenceTypes", "NONE"),
                            ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY"),
                            ref("inspectionQueryTypes", "GET_PACKAGE_SECURITY")),
                        List.of("A no-persistence audit pass reduces accidental mutation.")),
                    phase(
                        TaskPhasePurpose.ANALYZE,
                        "Analyze The Workbook",
                        "Read formula surfaces and aggregate workbook findings into one operator view.",
                        List.of(
                            ref("inspectionQueryTypes", "GET_FORMULA_SURFACE"),
                            ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS")),
                        List.of(
                            "Use targeted factual reads first when a later finding needs detail."))),
                List.of(
                    "EVENT_READ is intentionally limited; many audit plans require FULL_XSSF.",
                    "Loaded formulas that Apache POI cannot evaluate surface as UNSUPPORTED_FORMULA instead of silently recalculating.",
                    "If the audit needs strict policy checks, promote the findings into ASSERTION steps in a follow-up plan."))),
        starterPlan());
  }

  private static WorkbookPlan starterPlan() {
    String taskId = "AUDIT_EXISTING_WORKBOOK";
    String sourceAsset = TaskStarterRecipeSupport.taskStarterAsset("workbook-ops-source.xlsx");
    return ExampleWorkbookPlans.defaultExecutionPlan(
        TaskStarterRecipeSupport.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.ExistingFile(sourceAsset),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExampleSteps.read(
            "read-workbook-summary",
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary()),
        ExampleSteps.read(
            "read-template-comments",
            ExampleSelectors.cells("Template", "A1"),
            new SheetIntrospectionQuery.GetComments()),
        ExampleSteps.read(
            "read-drawing-objects",
            new DrawingObjectSelector.AllOnSheet("Template"),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
        ExampleSteps.read(
            "read-formula-surface",
            ExampleSelectors.sheets("Summary"),
            new InspectionSurfaceQuery.GetFormulaSurface()),
        ExampleSteps.read(
            "analyze-workbook-findings",
            ExampleSelectors.workbook(),
            new InspectionAnalysisQuery.AnalyzeWorkbookFindings()));
  }
}
