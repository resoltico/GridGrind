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
import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import java.util.List;
import java.util.Optional;

/** Published task recipe for intake and data-entry worksheet workflows. */
final class DataEntryWorkflowTaskRecipe {
  private DataEntryWorkflowTaskRecipe() {}

  static GridGrindTaskRecipeDefinition definition() {
    return GridGrindTaskRecipeSupport.definition(
        GridGrindTaskRecipeSupport.task(
            "DATA_ENTRY_WORKFLOW",
            GridGrindTaskRecipeSupport.selfContainedStarter("DATA_ENTRY_WORKFLOW"),
            List.of("office", "data entry", "validation", "intake", "worksheet"),
            GridGrindTaskRecipeSupport.discovery(
                List.of("data entry", "intake", "form", "worksheet", "validated input"),
                GridGrindTaskRecipeSupport.intent(
                    List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.SHEET,
                        TaskArtifactKind.DATA_VALIDATION,
                        TaskArtifactKind.COMMENT,
                        TaskArtifactKind.PROTECTION))),
            GridGrindTaskRecipeSupport.narrative(
                "Build one repeatable intake worksheet for row-oriented data entry with validations, comments, and later factual inspection.",
                List.of(
                    "Operators get one workbook surface designed for repeated entry instead of ad hoc edits.",
                    "Validation rules, comments, and protection settings are part of the authored workflow.",
                    "The result can be inspected or asserted after authoring."),
                List.of(
                    "Target sheet structure and protected cells.",
                    "Validation ranges and allowed values.",
                    "Save target when the intake workbook must be persisted."),
                List.of(
                    "Sheet protection after authoring.",
                    "Comments or prompts that explain the allowed entry flow.",
                    "Assertions on the protected or validated workbook surface.")),
            GridGrindTaskRecipeSupport.profile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.SAVE_AS,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            GridGrindTaskRecipeSupport.signals(
                List.of(
                    TaskInputKind.TARGET_SHEET_NAMES,
                    TaskInputKind.VALIDATION_RULES,
                    TaskInputKind.PERSISTENCE_TARGET_PATH),
                List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.ASSERTION_CHECKS)),
            GridGrindTaskRecipeSupport.workflow(
                List.of(
                    GridGrindTaskRecipeSupport.phase(
                        TaskPhasePurpose.PREPARE,
                        "Prepare The Intake Sheet",
                        "Create the sheet skeleton, labels, and writable cells first.",
                        List.of(
                            GridGrindTaskRecipeSupport.ref("sourceTypes", "NEW"),
                            GridGrindTaskRecipeSupport.ref("mutationActionTypes", "ENSURE_SHEET"),
                            GridGrindTaskRecipeSupport.ref("mutationActionTypes", "SET_RANGE"),
                            GridGrindTaskRecipeSupport.ref("mutationActionTypes", "SET_CELL")),
                        List.of(
                            "Start with the worksheet shape before layering validations or protection.")),
                    GridGrindTaskRecipeSupport.phase(
                        TaskPhasePurpose.AUTHOR,
                        "Author Guardrails",
                        "Add validations, comments, prompts, and protection that shape entry behavior.",
                        List.of(
                            GridGrindTaskRecipeSupport.ref(
                                "mutationActionTypes", "SET_DATA_VALIDATION"),
                            GridGrindTaskRecipeSupport.ref("mutationActionTypes", "SET_COMMENT"),
                            GridGrindTaskRecipeSupport.ref(
                                "mutationActionTypes", "SET_WORKBOOK_PROTECTION"),
                            GridGrindTaskRecipeSupport.ref("persistenceTypes", "SAVE_AS")),
                        List.of(
                            "Validation and protection belong to the authored model, not a memo.")),
                    GridGrindTaskRecipeSupport.phase(
                        TaskPhasePurpose.VERIFY,
                        "Inspect The Surface",
                        "Read back validations, comments, and summary facts before shipping the workbook.",
                        List.of(
                            GridGrindTaskRecipeSupport.ref(
                                "inspectionQueryTypes", "GET_DATA_VALIDATIONS"),
                            GridGrindTaskRecipeSupport.ref("inspectionQueryTypes", "GET_COMMENTS"),
                            GridGrindTaskRecipeSupport.ref(
                                "inspectionQueryTypes", "GET_WORKBOOK_SUMMARY")),
                        List.of(
                            "Factual rereads catch drift before the workbook reaches operators."))),
                List.of(
                    "Protection settings can restrict later mutation flows unless you plan for them.",
                    "Validation formulas and lists must fit the supported POI-backed contract shape.",
                    "Large instructional text belongs in external text sources instead of huge inline literals."))),
        starterPlan());
  }

  private static WorkbookPlan starterPlan() {
    String taskId = "DATA_ENTRY_WORKFLOW";
    return ExampleWorkbookPlans.defaultExecutionPlan(
        TaskStarterRecipeSupport.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(TaskStarterRecipeSupport.taskWorkbookPath(taskId)),
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
                    new DataValidationRuleInput.ExplicitList(List.of("Open", "Review", "Closed")),
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
            new WorkbookIntrospectionQuery.GetWorkbookSummary()));
  }
}
