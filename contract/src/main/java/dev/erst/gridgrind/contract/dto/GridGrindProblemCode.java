package dev.erst.gridgrind.contract.dto;

import java.util.Objects;
import java.util.Optional;

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
  ASSERTION_FAILED(
      GridGrindProblemCategory.ASSERTION,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Assertion failed",
      "Inspect the observed workbook facts, then adjust the plan expectations or authored"
          + " mutations and retry."),
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
  UNSUPPORTED_FORMULA_CONSTRUCT(
      GridGrindProblemCategory.FORMULA,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Unsupported formula construct",
      "Use a formula construct that Apache POI can parse on the authored request path, then"
          + " retry."),
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
      "Provide an existing authored input file or correct its path. When the CLI reads a request"
          + " via --request, relative request-owned paths resolve from the request file"
          + " directory."),
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
  WORKBOOK_PASSWORD_REQUIRED(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Workbook password required",
      "Provide source.security.password for the encrypted workbook and retry."),
  INVALID_WORKBOOK_PASSWORD(
      GridGrindProblemCategory.SECURITY,
      GridGrindProblemRecovery.CHANGE_REQUEST,
      "Invalid workbook password",
      "Supply the correct source.security.password for the encrypted workbook and retry."),
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
      case INVALID_ARGUMENTS -> invalidArgumentsResolution(message, context);
      case ASSERTION_FAILED ->
          "Inspect problem.assertionFailure observations, then adjust the failing assertion or"
              + " preceding workbook mutations and retry.";
      case IO_ERROR -> ioResolution(message, context);
      default -> resolution;
    };
  }

  private String invalidArgumentsResolution(String message, ProblemContext context) {
    String normalized = Objects.requireNonNullElse(message, "").trim();
    if (!(context instanceof ProblemContext.ParseArguments parseArguments)) {
      return resolution;
    }
    return specificInvalidArgumentsResolution(normalized)
        .orElseGet(
            () -> parseArgumentResolution(parseArguments.argumentName().orElse(""), normalized));
  }

  private static Optional<String> specificInvalidArgumentsResolution(String normalized) {
    if (normalized.startsWith("Unknown argument: ")) {
      return Optional.of(
          "Use one exact CLI flag. Start from --help for the synopsis, --help-protocol for the"
              + " grammar, or --help-guidance for workflow-oriented commands.");
    }
    if (normalized.startsWith("Ambiguous lookup id: ")) {
      return Optional.of(
          "Rerun the lookup with one qualified id exactly as listed in suggestions.");
    }
    if (normalized.startsWith("Unknown lookup id: ")) {
      return Optional.of(
          "Use --search when you know the concept but not the exact lookup id or group.");
    }
    if (normalized.startsWith("Unknown recipe: ")) {
      return Optional.of(
          "Use --print-recipe-catalog first when you need the stable recipe ids,"
              + " requestFileName, workspaceMode, and requiredWorkspacePaths, or"
              + " use --print-recipe-keyword-match when you know the goal but not the id.");
    }
    if (normalized.startsWith("No request JSON was provided.")) {
      return Optional.of(noRequestJsonResolution(normalized));
    }
    return Optional.empty();
  }

  private static String noRequestJsonResolution(String normalized) {
    if (normalized.contains("alongside --execution-root <path>")) {
      return "Use one real request document. Standard-input request mode always requires"
          + " --execution-root so relative request-owned paths resolve from one explicit"
          + " directory.";
    }
    return "Use --doctor-request only after you have one real request document to inspect.";
  }

  private static String parseArgumentResolution(String argumentName, String normalized) {
    return switch (argumentName) {
      case "--request" -> {
        if (normalized.contains("STANDARD_INPUT-authored payloads")) {
          yield "Requests that bind STANDARD_INPUT-authored payloads must arrive by --request so"
              + " standard input stays available for payload bytes.";
        }
        yield "Provide one readable request JSON file path, or omit --request and pipe one"
            + " request document on standard input.";
      }
      case "--execution-root" ->
          "When the request JSON arrives on stdin, pass one explicit --execution-root so"
              + " relative request-owned paths and internal temp files resolve from one"
              + " caller-chosen directory.";
      case "--response" -> "Provide one writable response file path after --response.";
      case "--lookup" ->
          "Use --lookup only with --print-recipe, --print-recipe-catalog, or"
              + " --print-protocol-catalog.";
      case "--query" ->
          normalized.startsWith("Invalid keyword query")
              ? "Use a natural-language query that leaves at least one searchable non-stop-word"
                  + " term after normalization."
              : "Use --query only with --print-recipe-keyword-match and provide one natural-language"
                  + " query.";
      case "--search" ->
          "Use --search only with --print-protocol-catalog and provide one search string.";
      default ->
          "Run gridgrind --help for the synopsis, --help-protocol for the authoritative request"
              + " contract, or --help-guidance for workflows and examples.";
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
                  ? "Choose a new --response path or remove the conflicting file, then retry."
                  : "Check the --response destination path, parent directory permissions, free"
                      + " disk space, and file locks before retrying."
              : resolution;
      default -> resolution;
    };
  }
}
