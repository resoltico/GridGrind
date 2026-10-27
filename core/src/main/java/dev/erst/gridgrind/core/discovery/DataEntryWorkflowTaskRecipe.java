package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import java.util.Optional;

/** Published task recipe for intake and data-entry worksheet workflows. */
final class DataEntryWorkflowTaskRecipe {
  private DataEntryWorkflowTaskRecipe() {}

  static TaskRecipeTemplate definition() {
    return TaskRecipes.definition(
        TaskRecipes.task(
            "DATA_ENTRY_WORKFLOW",
            TaskRecipes.selfContainedStarter("DATA_ENTRY_WORKFLOW"),
            List.of("office", "data entry", "validation", "intake", "worksheet"),
            TaskRecipes.discovery(
                List.of("data entry", "intake", "form", "worksheet", "validated input"),
                TaskRecipes.intent(
                    List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.SHEET,
                        TaskArtifactKind.DATA_VALIDATION,
                        TaskArtifactKind.COMMENT,
                        TaskArtifactKind.PROTECTION))),
            TaskRecipes.narrative(
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
            TaskRecipes.profile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.SAVE_AS,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            TaskRecipes.signals(
                List.of(
                    TaskInputKind.TARGET_SHEET_NAMES,
                    TaskInputKind.VALIDATION_RULES,
                    TaskInputKind.PERSISTENCE_TARGET_PATH),
                List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.ASSERTION_CHECKS)),
            TaskRecipes.workflow(
                List.of(
                    TaskRecipes.phase(
                        TaskPhasePurpose.PREPARE,
                        "Prepare The Intake Sheet",
                        "Create the sheet skeleton, labels, and writable cells first.",
                        List.of(
                            TaskRecipes.ref("workbookSourceTypes", "NEW"),
                            TaskRecipes.ref("mutationActionTypes", "ENSURE_SHEET"),
                            TaskRecipes.ref("mutationActionTypes", "SET_RANGE"),
                            TaskRecipes.ref("mutationActionTypes", "SET_CELL")),
                        List.of(
                            "Start with the worksheet shape before layering validations or protection.")),
                    TaskRecipes.phase(
                        TaskPhasePurpose.AUTHOR,
                        "Author Guardrails",
                        "Add validations, comments, prompts, and protection that shape entry behavior.",
                        List.of(
                            TaskRecipes.ref("mutationActionTypes", "SET_DATA_VALIDATION"),
                            TaskRecipes.ref("mutationActionTypes", "SET_COMMENT"),
                            TaskRecipes.ref("mutationActionTypes", "SET_WORKBOOK_PROTECTION"),
                            TaskRecipes.ref("workbookPersistenceTypes", "SAVE_AS")),
                        List.of(
                            "Validation and protection belong to the authored model, not a memo.")),
                    TaskRecipes.phase(
                        TaskPhasePurpose.VERIFY,
                        "Inspect The Surface",
                        "Read back validations, comments, and summary facts before shipping the workbook.",
                        List.of(
                            TaskRecipes.ref("inspectionQueryTypes", "GET_DATA_VALIDATIONS"),
                            TaskRecipes.ref("inspectionQueryTypes", "GET_COMMENTS"),
                            TaskRecipes.ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY")),
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
        TaskStarters.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(TaskStarters.taskWorkbookPath(taskId)),
        ExampleSteps.step(
            new StepId("ensure-intake"),
            ExampleSelectors.sheet("Intake"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("seed-intake-headers"),
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
            new StepId("comment-status-header"),
            ExampleSelectors.cell("Intake", "B1"),
            new CellMutationAction.SetComment(
                CommentSpec.plain(
                    TextSourceSpec.inline("Allowed values are Open, Review, or Closed."),
                    "GridGrind",
                    true))),
        ExampleSteps.step(
            new StepId("validate-status"),
            ExampleSelectors.range("Intake", "B2:B25"),
            new StructuredMutationAction.SetDataValidation(
                new DataValidationSpec(
                    new DataValidationRuleSpec.ExplicitList(List.of("Open", "Review", "Closed")),
                    false,
                    false,
                    Optional.of(
                        new DataValidationPromptSpec(
                            TextSourceSpec.inline("Status"),
                            TextSourceSpec.inline("Choose one approved workflow state."),
                            true)),
                    Optional.of(
                        new DataValidationErrorAlertSpec(
                            ExcelDataValidationErrorStyle.STOP,
                            TextSourceSpec.inline("Invalid status"),
                            TextSourceSpec.inline("Use Open, Review, or Closed."),
                            true))))),
        ExampleSteps.step(
            new StepId("protect-workbook-structure"),
            ExampleSelectors.workbook(),
            new WorkbookMutationAction.SetWorkbookProtection(
                new WorkbookProtectionSpec(
                    true, false, false, Optional.empty(), Optional.empty()))),
        ExampleSteps.read(
            new StepId("read-intake-validations"),
            ExampleSelectors.range("Intake", "B2:B25"),
            new SheetIntrospectionQuery.GetDataValidations()),
        ExampleSteps.read(
            new StepId("read-intake-comments"),
            ExampleSelectors.cells("Intake", "B1"),
            new SheetIntrospectionQuery.GetComments()),
        ExampleSteps.read(
            new StepId("read-intake-workbook"),
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary()));
  }
}
