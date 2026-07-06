package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.TaskArtifactKind;
import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskGoalKind;
import dev.erst.gridgrind.cli.discovery.TaskInputKind;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.cli.discovery.TaskVerificationKind;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.List;
import java.util.Optional;

/** Published task recipe for drawing and signature-line workflows. */
final class DrawingAndSignatureWorkflowTaskRecipe {
  private DrawingAndSignatureWorkflowTaskRecipe() {}

  static GridGrindTaskRecipeDefinition definition() {
    return GridGrindTaskRecipeSupport.definition(
        GridGrindTaskRecipeSupport.task(
            "DRAWING_AND_SIGNATURE_WORKFLOW",
            GridGrindTaskRecipeSupport.selfContainedStarter("DRAWING_AND_SIGNATURE_WORKFLOW"),
            List.of("office", "drawing", "signature", "picture", "embedded-object"),
            GridGrindTaskRecipeSupport.discovery(
                List.of("drawing", "signature line", "picture", "shape", "embedded object"),
                GridGrindTaskRecipeSupport.intent(
                    List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.DRAWING_OBJECT,
                        TaskArtifactKind.SIGNATURE_LINE))),
            GridGrindTaskRecipeSupport.narrative(
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
            GridGrindTaskRecipeSupport.profile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.SAVE_AS,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            GridGrindTaskRecipeSupport.signals(
                List.of(
                    TaskInputKind.TARGET_SHEET_NAMES,
                    TaskInputKind.TARGET_OBJECT_NAMES,
                    TaskInputKind.DRAWING_ANCHORS,
                    TaskInputKind.PERSISTENCE_TARGET_PATH),
                List.of(TaskVerificationKind.FACT_READBACK)),
            GridGrindTaskRecipeSupport.workflow(
                List.of(
                    GridGrindTaskRecipeSupport.phase(
                        TaskPhasePurpose.PREPARE,
                        "Prepare The Drawing Surface",
                        "Create the target sheet before any named drawing object is authored.",
                        List.of(
                            GridGrindTaskRecipeSupport.ref("sourceTypes", "NEW"),
                            GridGrindTaskRecipeSupport.ref("mutationActionTypes", "ENSURE_SHEET")),
                        List.of("A stable sheet name gives drawing objects a durable home.")),
                    GridGrindTaskRecipeSupport.phase(
                        TaskPhasePurpose.AUTHOR,
                        "Author Drawing Objects",
                        "Create one named drawing object and move it authoritatively when needed.",
                        List.of(
                            GridGrindTaskRecipeSupport.ref(
                                "mutationActionTypes", "SET_SIGNATURE_LINE"),
                            GridGrindTaskRecipeSupport.ref(
                                "mutationActionTypes", "SET_DRAWING_OBJECT_ANCHOR"),
                            GridGrindTaskRecipeSupport.ref("persistenceTypes", "SAVE_AS")),
                        List.of("Named drawing objects are easier to inspect and move later.")),
                    GridGrindTaskRecipeSupport.phase(
                        TaskPhasePurpose.VERIFY,
                        "Inspect The Result",
                        "Read back worksheet drawing metadata after authoring or anchor changes.",
                        List.of(
                            GridGrindTaskRecipeSupport.ref(
                                "inspectionQueryTypes", "GET_DRAWING_OBJECTS")),
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
        TaskStarterRecipeSupport.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(TaskStarterRecipeSupport.taskWorkbookPath(taskId)),
        ExampleSteps.step(
            "ensure-approvals",
            ExampleSelectors.sheet("Approvals"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            "set-signature-line",
            ExampleSelectors.sheet("Approvals"),
            new DrawingMutationAction.SetSignatureLine(
                new SignatureLineInput(
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
                        new PictureDataInput(
                            ExcelPictureFormat.PNG,
                            BinarySourceInput.inlineBase64(
                                TaskStarterRecipeSupport.onePixelPngBase64())))))),
        ExampleSteps.read(
            "read-drawing-objects",
            new DrawingObjectSelector.AllOnSheet("Approvals"),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
        ExampleSteps.step(
            "move-signature-line",
            new DrawingObjectSelector.ByName("Approvals", "WorkflowSignature"),
            new DrawingMutationAction.SetDrawingObjectAnchor(
                ExampleDrawingAnchors.anchor(5, 1, 8, 6))),
        ExampleSteps.read(
            "read-drawing-objects-after-move",
            new DrawingObjectSelector.AllOnSheet("Approvals"),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjects()));
  }
}
