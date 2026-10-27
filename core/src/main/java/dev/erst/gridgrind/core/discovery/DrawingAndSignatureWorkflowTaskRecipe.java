package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import java.util.Optional;

/** Published task recipe for drawing and signature-line workflows. */
final class DrawingAndSignatureWorkflowTaskRecipe {
  private DrawingAndSignatureWorkflowTaskRecipe() {}

  static TaskRecipeTemplate definition() {
    return TaskRecipes.definition(
        TaskRecipes.task(
            "DRAWING_AND_SIGNATURE_WORKFLOW",
            TaskRecipes.selfContainedStarter("DRAWING_AND_SIGNATURE_WORKFLOW"),
            List.of("office", "drawing", "signature", "picture", "embedded-object"),
            TaskRecipes.discovery(
                List.of("drawing", "signature line", "picture", "shape", "embedded object"),
                TaskRecipes.intent(
                    List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.DRAWING_OBJECT,
                        TaskArtifactKind.SIGNATURE_LINE))),
            TaskRecipes.narrative(
                "Author or inspect drawing-backed workbook content such as signature lines, pictures, shapes, and embedded objects.",
                List.of(
                    "Drawing-backed workbook objects are authored by name instead of fragile XML surgery.",
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
            TaskRecipes.profile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.SAVE_AS,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            TaskRecipes.signals(
                List.of(
                    TaskInputKind.TARGET_SHEET_NAMES,
                    TaskInputKind.TARGET_OBJECT_NAMES,
                    TaskInputKind.DRAWING_ANCHORS,
                    TaskInputKind.PERSISTENCE_TARGET_PATH),
                List.of(TaskVerificationKind.FACT_READBACK)),
            TaskRecipes.workflow(
                List.of(
                    TaskRecipes.phase(
                        TaskPhasePurpose.PREPARE,
                        "Prepare The Drawing Surface",
                        "Create the target sheet before any named drawing object is authored.",
                        List.of(
                            TaskRecipes.ref("workbookSourceTypes", "NEW"),
                            TaskRecipes.ref("mutationActionTypes", "ENSURE_SHEET")),
                        List.of("A stable sheet name gives drawing objects a durable home.")),
                    TaskRecipes.phase(
                        TaskPhasePurpose.AUTHOR,
                        "Author Drawing Objects",
                        "Create one named drawing object and move it authoritatively when needed.",
                        List.of(
                            TaskRecipes.ref("mutationActionTypes", "SET_SIGNATURE_LINE"),
                            TaskRecipes.ref("mutationActionTypes", "SET_DRAWING_OBJECT_ANCHOR"),
                            TaskRecipes.ref("workbookPersistenceTypes", "SAVE_AS")),
                        List.of("Named drawing objects are easier to inspect and move later.")),
                    TaskRecipes.phase(
                        TaskPhasePurpose.VERIFY,
                        "Inspect The Result",
                        "Read back worksheet drawing metadata after authoring or anchor changes.",
                        List.of(TaskRecipes.ref("inspectionQueryTypes", "GET_DRAWING_OBJECTS")),
                        List.of("Use factual reread to confirm the visible workbook surface."))),
                List.of(
                    "Signature lines are VML-backed drawing objects with workbook-side constraints.",
                    "Invalid drawing payloads are rejected without leaving partial artifacts behind.",
                    "Image-heavy workflows should prefer file-backed binary sources over huge inline base64 payloads."))),
        starterPlan());
  }

  private static WorkbookPlan starterPlan() {
    String taskId = "DRAWING_AND_SIGNATURE_WORKFLOW";
    return ExampleWorkbookPlans.defaultExecutionPlan(
        TaskStarters.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(TaskStarters.taskWorkbookPath(taskId)),
        ExampleSteps.step(
            new StepId("ensure-approvals"),
            ExampleSelectors.sheet("Approvals"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("set-signature-line"),
            ExampleSelectors.sheet("Approvals"),
            new DrawingMutationAction.SetSignatureLine(
                new SignatureLineSpec(
                    "WorkflowSignature",
                    ExampleDrawingAnchors.anchor(1, 1, 4, 6),
                    false,
                    Optional.of("Review the workflow before signing."),
                    Optional.of("Ada Lovelace"),
                    Optional.of("Operations"),
                    Optional.of("ada@example.com"),
                    Optional.empty(),
                    Optional.of("invalid"),
                    Optional.of(
                        new PictureDataSpec(
                            ExcelPictureFormat.PNG,
                            BinarySourceSpec.inlineBase64(TaskStarters.onePixelPngBase64())))))),
        ExampleSteps.read(
            new StepId("read-drawing-objects"),
            new DrawingObjectSelector.AllOnSheet(new SheetName("Approvals")),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
        ExampleSteps.step(
            new StepId("move-signature-line"),
            new DrawingObjectSelector.ByName(new SheetName("Approvals"), "WorkflowSignature"),
            new DrawingMutationAction.SetDrawingObjectAnchor(
                ExampleDrawingAnchors.anchor(5, 1, 8, 6))),
        ExampleSteps.read(
            new StepId("read-drawing-objects-after-move"),
            new DrawingObjectSelector.AllOnSheet(new SheetName("Approvals")),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjects()));
  }
}
