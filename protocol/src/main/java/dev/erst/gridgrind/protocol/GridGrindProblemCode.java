package dev.erst.gridgrind.protocol;

/** Stable machine-readable problem codes returned by the agent protocol. */
public enum GridGrindProblemCode {
  INVALID_ARGUMENTS(
      GridGrindProblemCategory.ARGUMENTS,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid CLI arguments",
      "Fix the CLI arguments and rerun the command."),
  INVALID_JSON(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid JSON payload",
      "Send syntactically valid JSON."),
  INVALID_REQUEST_SHAPE(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid request shape",
      "Send a payload whose fields and discriminator IDs match the GridGrind protocol."),
  INVALID_REQUEST(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid request",
      "Fix the request data and retry the workflow."),
  INVALID_CELL_ADDRESS(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid cell address",
      "Use a valid A1-style address such as A1 or BC12."),
  INVALID_RANGE_ADDRESS(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid range address",
      "Use a valid A1-style range such as A1:C3 or a single address like B2."),
  INVALID_FORMULA(
      GridGrindProblemCategory.FORMULA,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid formula",
      "Fix the formula syntax or workbook references, then retry."),
  UNSUPPORTED_FORMULA(
      GridGrindProblemCategory.FORMULA,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Unsupported formula",
      "Use a formula supported by Apache POI or skip server-side formula evaluation."),
  WORKBOOK_NOT_FOUND(
      GridGrindProblemCategory.RESOURCE,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Workbook not found",
      "Create the workbook first or provide an existing workbook path."),
  SHEET_NOT_FOUND(
      GridGrindProblemCategory.RESOURCE,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Sheet not found",
      "Create the sheet first or correct the sheet name."),
  NAMED_RANGE_NOT_FOUND(
      GridGrindProblemCategory.RESOURCE,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Named range not found",
      "Create the named range first or correct the requested scope and name."),
  CELL_NOT_FOUND(
      GridGrindProblemCategory.RESOURCE,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Cell not found",
      "Write the cell first or adjust the analysis target."),
  IO_ERROR(
      GridGrindProblemCategory.IO,
      GridGrindProblemRecovery.CHECK_ENVIRONMENT,
      "I/O failure",
      "Check file paths, permissions, file locks, and disk state before retrying."),
  INTERNAL_ERROR(
      GridGrindProblemCategory.INTERNAL,
      GridGrindProblemRecovery.ESCALATE,
      "Internal GridGrind failure",
      "Capture the problem details and escalate; this indicates an unexpected runtime failure.");

  private final GridGrindProblemCategory category;
  private final GridGrindProblemRecovery recovery;
  private final String title;
  private final String resolution;

  GridGrindProblemCode(
      GridGrindProblemCategory category,
      GridGrindProblemRecovery recovery,
      String title,
      String resolution) {
    this.category = category;
    this.recovery = recovery;
    this.title = title;
    this.resolution = resolution;
  }

  public GridGrindProblemCategory category() {
    return category;
  }

  public GridGrindProblemRecovery recovery() {
    return recovery;
  }

  public String title() {
    return title;
  }

  public String resolution() {
    return resolution;
  }
}
