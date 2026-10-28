package dev.erst.gridgrind.core.problem;

import java.util.Objects;

/** Stable machine-readable problem codes returned by the agent protocol. */
public enum GridGrindProblemCode {
  INVALID_ARGUMENTS(
      GridGrindProblemCategory.ARGUMENTS,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid host arguments",
      "Fix the host arguments and rerun the command."),
  INVALID_JSON(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid JSON payload",
      "Send syntactically valid JSON."),
  INVALID_ENCODING(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid request encoding",
      "Encode the request document as UTF-8."),
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
  NUMBER_NOT_REPRESENTABLE(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Number cannot be represented exactly",
      "Store identifiers or precision-sensitive values as TEXT, or use a value exactly"
          + " representable as an IEEE-754 number."),
  PATH_ESCAPES_ROOT(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Path escapes execution root",
      "Use a request-owned path inside the execution root; absolute paths are allowed only when"
          + " they resolve inside that root."),
  ASSERTION_FAILED(
      GridGrindProblemCategory.ASSERTION,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Assertion failed",
      "Inspect the observed workbook facts, then adjust the plan expectations or authored"
          + " mutations and retry."),
  PRESERVATION_FAILED(
      GridGrindProblemCategory.ASSERTION,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Preservation requirement failed",
      "Inspect the protected facts, then adjust the plan mutations or the host-required"
          + " preservation selection and retry."),
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
  INVALID_FORMULA_TEXT(
      GridGrindProblemCategory.FORMULA,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid opaque formula text",
      "Provide nonempty XML 1.0 formula character data without a leading =."),
  UNSUPPORTED_FORMULA_CONSTRUCT(
      GridGrindProblemCategory.FORMULA,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Unsupported formula construct",
      "Use a formula construct that Apache POI can parse on the authored request path, then"
          + " retry."),
  CIRCULAR_FORMULA_REFERENCE(
      GridGrindProblemCategory.FORMULA,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Circular formula reference",
      "Break the circular formula dependency or use a non-evaluating calculation strategy."),
  MISSING_EXTERNAL_WORKBOOK(
      GridGrindProblemCategory.FORMULA,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Missing external workbook",
      "Bind the referenced external workbook or switch to cached-value evaluation."),
  UNREGISTERED_USER_DEFINED_FUNCTION(
      GridGrindProblemCategory.FORMULA,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Unregistered user-defined function",
      "Register the required UDF in formulaEnvironment or skip server-side evaluation."),
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
  INPUT_SOURCE_NOT_FOUND(
      GridGrindProblemCategory.RESOURCE,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Input source not found",
      "Provide an existing authored input file or correct its path. Relative request-owned"
          + " paths resolve from the explicit execution root."),
  INPUT_SOURCE_UNAVAILABLE(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Input source unavailable",
      "Bind the required STANDARD_INPUT content or change the request to use INLINE or FILE"
          + " sources."),
  INPUT_SOURCE_IO_ERROR(
      GridGrindProblemCategory.IO,
      GridGrindProblemRecovery.CHECK_ENVIRONMENT,
      "Input source I/O failure",
      "Check authored input paths, permissions, file locks, and transport bindings before"
          + " retrying."),
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
  WORKBOOK_NOT_OPENABLE(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Workbook package is not openable",
      "Provide a valid `.xlsx` OOXML workbook package."),
  SOURCE_PATH_IS_DIRECTORY(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Workbook source path is a directory",
      "Provide a workbook file path instead of a directory."),
  OUTPUT_PATH_IS_DIRECTORY(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Workbook output path is a directory",
      "Provide a workbook file path instead of a directory."),
  OUTPUT_PATH_ALREADY_EXISTS(
      GridGrindProblemCategory.REQUEST,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Workbook output path already exists",
      "Choose a new output path, use ifExists=REPLACE, or remove the conflicting file."),
  WORKBOOK_PASSWORD_REQUIRED(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Workbook password required",
      "Provide source.security.passwordRef backed by an authorized host secret, then retry."),
  INVALID_WORKBOOK_PASSWORD(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid workbook password",
      "Check the secret resolved by source.security.passwordRef, then retry."),
  ENCRYPTION_SOURCE_NOT_ENCRYPTED(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Source workbook is not encrypted",
      "Use persistence.security.encryption.type=NONE or ENCRYPT for a plaintext source workbook."),
  ENCRYPTION_SOURCE_NOT_PRESERVABLE(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Source encryption cannot be preserved",
      "Use persistence.security.encryption.type=NONE or ENCRYPT when the source encryption"
          + " envelope is not compatible with GridGrind's AGILE write contract."),
  INVALID_SIGNING_CONFIGURATION(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid signing configuration",
      "Fix persistence.security.signature keystore, alias, password, or digest settings and"
          + " retry."),
  WORKBOOK_SECURITY_ERROR(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHECK_ENVIRONMENT,
      "Workbook security failure",
      "Check the secure workbook package, cryptographic material, and runtime environment before"
          + " retrying."),
  UNSAFE_PATH_ACCESS(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHECK_ENVIRONMENT,
      "Unsafe path access",
      "Use a filesystem that supports GridGrind's no-follow path binding and retry after any"
          + " observed path topology change is resolved."),
  AUTHORITY_DENIED(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Execution authority denied",
      "Request a host grant that explicitly permits the required resource, operation, target, or"
          + " publication effect."),
  PUBLICATION_UNCERTAIN(
      GridGrindProblemCategory.IO,
      GridGrindProblemRecovery.ESCALATE,
      "Workbook publication state is uncertain",
      "Inspect the destination workbook and its recorded publication evidence before any retry;"
          + " do not rerun the plan blindly."),
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

  /** Returns the most specific public remediation text for the classified problem cause. */
  public String resolutionFor(String message, ProblemContext context) {
    Objects.requireNonNull(context, "context must not be null");
    return switch (this) {
      case INVALID_ARGUMENTS -> resolution;
      case ASSERTION_FAILED ->
          "Inspect problem.assertionFailure observations, then adjust the failing assertion or"
              + " preceding workbook mutations and retry.";
      case IO_ERROR -> ioResolution(message, context);
      default -> resolution;
    };
  }

  private String ioResolution(String message, ProblemContext context) {
    String normalized = Objects.requireNonNullElse(message, "").trim();
    return switch (context) {
      case ProblemContext.PersistWorkbook persistWorkbook -> {
        if (persistWorkbook.persistencePath().isPresent()) {
          yield normalized.contains("already exists")
              ? "Choose a new SAVE_AS destination path, remove the conflicting file, or set"
                  + " SAVE_AS.ifExists=REPLACE before retrying."
              : "Check the SAVE_AS destination path, parent directory permissions, free disk"
                  + " space, and file locks before retrying.";
        }
        yield "Check the destination workbook path, permissions, free disk space, and file"
            + " locks before retrying.";
      }
      case ProblemContext.OpenWorkbook openWorkbook ->
          openWorkbook.sourceWorkbookPath().isPresent()
              ? "Check the source workbook path, permissions, and file locks before retrying."
              : resolution;
      case ProblemContext.WriteResponse writeResponse ->
          writeResponse.responsePath().isPresent()
              ? normalized.contains("already exists")
                  ? "Choose a new response destination or remove the conflicting file, then retry."
                  : "Check the response destination path, parent directory permissions, free"
                      + " disk space, and file locks before retrying."
              : resolution;
      default -> resolution;
    };
  }
}
