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
        ExamplePlanSupport.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExamplePlanSupport.step(
                "ensure-intake",
                ExamplePlanSupport.sheet("Intake"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "seed-intake-headers",
                ExamplePlanSupport.range("Intake", "A1:B3"),
                new CellMutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Owner"), ExamplePlanSupport.text("Status")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Ada"), ExamplePlanSupport.text("Open")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Lin"), ExamplePlanSupport.text("Review"))))),
            ExamplePlanSupport.step(
                "comment-status-header",
                ExamplePlanSupport.cell("Intake", "B1"),
                new CellMutationAction.SetComment(
                    CommentInput.plain(
                        TextSourceInput.inline("Allowed values are Open, Review, or Closed."),
                        "GridGrind",
                        true))),
            ExamplePlanSupport.step(
                "validate-status",
                ExamplePlanSupport.range("Intake", "B2:B25"),
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
            ExamplePlanSupport.step(
                "protect-workbook-structure",
                ExamplePlanSupport.workbook(),
                new WorkbookMutationAction.SetWorkbookProtection(
                    new WorkbookProtectionInput(
                        true, false, false, Optional.empty(), Optional.empty()))),
            ExamplePlanSupport.read(
                "read-intake-validations",
                ExamplePlanSupport.range("Intake", "B2:B25"),
                new SheetIntrospectionQuery.GetDataValidations()),
            ExamplePlanSupport.read(
                "read-intake-comments",
                ExamplePlanSupport.cells("Intake", "B1"),
                new SheetIntrospectionQuery.GetComments()),
            ExamplePlanSupport.read(
                "read-intake-workbook",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary())));
  }

  private static TaskStarterPlan drawingStarter() {
    String taskId = "DRAWING_AND_SIGNATURE_WORKFLOW";
    return TaskStarterPlanSupport.selfContainedStarter(
        taskId,
        ExamplePlanSupport.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExamplePlanSupport.step(
                "ensure-approvals",
                ExamplePlanSupport.sheet("Approvals"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "set-signature-line",
                ExamplePlanSupport.sheet("Approvals"),
                new DrawingMutationAction.SetSignatureLine(
                    new SignatureLineInput(
                        "WorkflowSignature",
                        ExamplePlanSupport.anchor(1, 1, 4, 6),
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
            ExamplePlanSupport.read(
                "read-drawing-objects",
                new DrawingObjectSelector.AllOnSheet("Approvals"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            ExamplePlanSupport.step(
                "move-signature-line",
                new DrawingObjectSelector.ByName("Approvals", "WorkflowSignature"),
                new DrawingMutationAction.SetDrawingObjectAnchor(
                    ExamplePlanSupport.anchor(5, 1, 8, 6))),
            ExamplePlanSupport.read(
                "read-drawing-objects-after-move",
                new DrawingObjectSelector.AllOnSheet("Approvals"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects())));
  }

  private static TaskStarterPlan workbookMaintenanceStarter() {
    String taskId = "WORKBOOK_MAINTENANCE";
    String sourceAsset = TaskStarterPlanSupport.taskStarterAsset("workbook-ops-source.xlsx");
    return TaskStarterPlanSupport.assetBackedStarter(
        taskId,
        ExamplePlanSupport.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new WorkbookPlan.WorkbookSource.ExistingFile(sourceAsset),
            ExamplePlanSupport.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExamplePlanSupport.read(
                "read-source-workbook",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.step(
                "copy-template-sheet",
                ExamplePlanSupport.sheet("Template"),
                new WorkbookMutationAction.CopySheet("Template Copy")),
            ExamplePlanSupport.read(
                "read-copied-comments",
                new dev.erst.gridgrind.contract.selector.CellSelector.AllUsedInSheet(
                    "Template Copy"),
                new SheetIntrospectionQuery.GetComments()),
            ExamplePlanSupport.read(
                "read-copied-drawing-objects",
                new DrawingObjectSelector.AllOnSheet("Template Copy"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            ExamplePlanSupport.read(
                "read-maintenance-findings",
                ExamplePlanSupport.workbook(),
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings())),
        sourceAsset);
  }
}
