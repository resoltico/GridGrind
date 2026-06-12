package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.List;
import java.util.Optional;

/** Published starter plans for workflow authoring and workbook maintenance. */
final class GridGrindTaskStarterWorkflowPlans {
  private GridGrindTaskStarterWorkflowPlans() {}

  static List<TaskStarterPlan> starters() {
    return List.of(dataEntryStarter(), drawingStarter(), workbookMaintenanceStarter());
  }

  private static TaskStarterPlan dataEntryStarter() {
    String taskId = "DATA_ENTRY_WORKFLOW";
    return TaskStarterPlanSupport.selfContainedStarter(
        taskId,
        ExampleWorkbookPlans.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.New(),
            ExampleWorkbookPlans.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExampleSteps.step(
                "ensure-intake",
                ExampleSelectors.sheet("Intake"),
                new WorkbookMutationAction.EnsureSheet()),
            ExampleSteps.step(
                "seed-intake-headers",
                ExampleSelectors.range("Intake", "A1:B3"),
                new CellMutationAction.SetRange(
                    ExampleCellValues.rows(
                        ExampleCellValues.row(
                            ExampleCellValues.text("Owner"), ExampleCellValues.text("Status")),
                        ExampleCellValues.row(
                            ExampleCellValues.text("Ada"), ExampleCellValues.text("Open")),
                        ExampleCellValues.row(
                            ExampleCellValues.text("Lin"), ExampleCellValues.text("Review"))))),
            ExampleSteps.step(
                "comment-status-header",
                ExampleSelectors.cell("Intake", "B1"),
                new CellMutationAction.SetComment(
                    CommentInput.plain(
                        TextSourceInput.inline("Allowed values are Open, Review, or Closed."),
                        "GridGrind",
                        true))),
            ExampleSteps.step(
                "validate-status",
                ExampleSelectors.range("Intake", "B2:B25"),
                new StructuredMutationAction.SetDataValidation(
                    new DataValidationInput(
                        new DataValidationRuleInput.ExplicitList(
                            List.of("Open", "Review", "Closed")),
                        false,
                        false,
                        Optional.of(
                            new DataValidationPromptInput(
                                TextSourceInput.inline("Status"),
                                TextSourceInput.inline("Choose one approved workflow state."),
                                true)),
                        Optional.of(
                            new DataValidationErrorAlertInput(
                                ExcelDataValidationErrorStyle.STOP,
                                TextSourceInput.inline("Invalid status"),
                                TextSourceInput.inline("Use Open, Review, or Closed."),
                                true))))),
            ExampleSteps.step(
                "protect-workbook-structure",
                ExampleSelectors.workbook(),
                new WorkbookMutationAction.SetWorkbookProtection(
                    new WorkbookProtectionInput(
                        true, false, false, Optional.empty(), Optional.empty()))),
            ExampleSteps.read(
                "read-intake-validations",
                ExampleSelectors.range("Intake", "B2:B25"),
                new SheetIntrospectionQuery.GetDataValidations()),
            ExampleSteps.read(
                "read-intake-comments",
                ExampleSelectors.cells("Intake", "B1"),
                new SheetIntrospectionQuery.GetComments()),
            ExampleSteps.read(
                "read-intake-workbook",
                ExampleSelectors.workbook(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary())));
  }

  private static TaskStarterPlan drawingStarter() {
    String taskId = "DRAWING_AND_SIGNATURE_WORKFLOW";
    return TaskStarterPlanSupport.selfContainedStarter(
        taskId,
        ExampleWorkbookPlans.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.New(),
            ExampleWorkbookPlans.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
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
                                    TaskStarterPlanSupport.onePixelPngBase64())))))),
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
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects())));
  }

  private static TaskStarterPlan workbookMaintenanceStarter() {
    String taskId = "WORKBOOK_MAINTENANCE";
    String sourceAsset = TaskStarterPlanSupport.taskStarterAsset("workbook-ops-source.xlsx");
    return TaskStarterPlanSupport.assetBackedStarter(
        taskId,
        ExampleWorkbookPlans.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.ExistingFile(sourceAsset),
            ExampleWorkbookPlans.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
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
                new dev.erst.gridgrind.contract.selector.CellSelector.AllUsedInSheet(
                    "Template Copy"),
                new SheetIntrospectionQuery.GetComments()),
            ExampleSteps.read(
                "read-copied-drawing-objects",
                new DrawingObjectSelector.AllOnSheet("Template Copy"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            ExampleSteps.read(
                "read-maintenance-findings",
                ExampleSelectors.workbook(),
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings())),
        sourceAsset);
  }
}
