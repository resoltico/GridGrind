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

/** Task descriptor for intake/data-entry worksheet workflows. */
final class DataEntryWorkflowTaskDefinition {
  private DataEntryWorkflowTaskDefinition() {}

  static TaskEntry entry() {
    return task(
        "DATA_ENTRY_WORKFLOW",
        discovery(
            List.of("data entry", "intake", "form", "worksheet", "validated input"),
            List.of("office", "data entry", "validation", "intake", "worksheet"),
            intent(
                List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY),
                List.of(
                    TaskArtifactKind.WORKBOOK,
                    TaskArtifactKind.SHEET,
                    TaskArtifactKind.DATA_VALIDATION,
                    TaskArtifactKind.COMMENT,
                    TaskArtifactKind.PROTECTION))),
        narrative(
            "Build one repeatable intake worksheet for row-oriented data entry with validations,"
                + " comments, and later factual inspection.",
            List.of(
                "Operators get one workbook surface designed for repeated entry instead of ad hoc"
                    + " edits.",
                "Validation rules, comments, and protection settings are part of the authored"
                    + " workflow.",
                "The result can be inspected or asserted after authoring."),
            List.of(
                "Target sheet structure and protected cells.",
                "Validation ranges and allowed values.",
                "Save target when the intake workbook must be persisted."),
            List.of(
                "Sheet protection after authoring.",
                "Comments or prompts that explain the allowed entry flow.",
                "Assertions on the protected or validated workbook surface.")),
        profile(
            TaskSourceMode.NEW_WORKBOOK,
            TaskPersistenceMode.SAVE_AS,
            TaskMutationMode.MUTATING,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(
                TaskInputKind.TARGET_SHEET_NAMES,
                TaskInputKind.VALIDATION_RULES,
                TaskInputKind.PERSISTENCE_TARGET_PATH),
            List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.ASSERTION_CHECKS)),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.PREPARE,
                    "Prepare The Intake Sheet",
                    "Create the sheet skeleton, labels, and writable cells first.",
                    List.of(
                        ref("sourceTypes", "NEW"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("mutationActionTypes", "SET_RANGE"),
                        ref("mutationActionTypes", "SET_CELL")),
                    List.of(
                        "Start with the worksheet shape before layering validations or"
                            + " protection.")),
                phase(
                    TaskPhasePurpose.AUTHOR,
                    "Author Guardrails",
                    "Add validations, comments, prompts, and protection that shape entry"
                        + " behavior.",
                    List.of(
                        ref("mutationActionTypes", "SET_DATA_VALIDATION"),
                        ref("mutationActionTypes", "SET_COMMENT"),
                        ref("mutationActionTypes", "SET_WORKBOOK_PROTECTION"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Validation and protection belong to the authored model, not a memo.")),
                phase(
                    TaskPhasePurpose.VERIFY,
                    "Inspect The Surface",
                    "Read back validations, comments, and summary facts before shipping the"
                        + " workbook.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_DATA_VALIDATIONS"),
                        ref("inspectionQueryTypes", "GET_COMMENTS"),
                        ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY")),
                    List.of("Factual rereads catch drift before the workbook reaches operators."))),
            List.of(
                "Protection settings can restrict later mutation flows unless you plan for"
                    + " them.",
                "Validation formulas and lists must fit the supported POI-backed contract"
                    + " shape.",
                "Large instructional text belongs in external text sources instead of huge"
                    + " inline literals.")));
  }
}
