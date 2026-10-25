package dev.erst.gridgrind.contract.catalog;

/** Request-surface, CLI-path, and scratch-ownership wording shared by public contract surfaces. */
public final class GridGrindRequestSurfaceContractText {
  private static final long REQUEST_DOCUMENT_LIMIT_BYTES = 16L * 1024 * 1024;

  private GridGrindRequestSurfaceContractText() {}

  /** One stable help and runtime message for stdin-backed authored values. */
  public static String standardInputRequiresRequestMessage() {
    return "STANDARD_INPUT-authored values require --request so stdin is available for input"
        + " content instead of the request JSON";
  }

  /** One stable help and runtime message for stdin-rooted request execution. */
  public static String stdinExecutionRootRequiredMessage() {
    return "Requests read from stdin require --execution-root so request-owned paths resolve"
        + " from one explicit directory";
  }

  /** Stable wording for how relative paths inside the request are resolved. */
  public static String requestOwnedPathResolutionSummary() {
    return "When the CLI reads a request via --request, relative request-owned paths resolve from"
        + " the request file directory. When the request JSON arrives on stdin, pass"
        + " --execution-root <path> and relative request-owned paths resolve from that"
        + " directory.";
  }

  /** Stable wording for how CLI file-flag paths are resolved. */
  public static String cliFlagPathResolutionSummary() {
    return "--request and --response resolve from the current working directory, as do"
        + " --execution-root and --temp-root.";
  }

  /** Stable wording for how CLI-managed scratch space is owned and defaulted. */
  public static String cliScratchSpaceSummary() {
    return "GridGrind creates one private per-run scratch directory. Without --temp-root, that"
        + " scratch directory lives under the OS temporary-file root. With --temp-root <path>,"
        + " GridGrind creates the private scratch directory under that supplied parent path."
        + " Best-effort cleanup removes the managed scratch directory on normal command"
        + " completion.";
  }

  /** Stable wording for encrypted OOXML plaintext temp ownership. */
  public static String encryptedWorkbookTempSecuritySummary() {
    return "Encrypted OOXML plaintext temp workbooks always stay in one private OS temporary"
        + " directory rather than the request root, execution root, or CLI scratch-parent"
        + " override.";
  }

  /** Maximum accepted JSON request document size in bytes. */
  public static long requestDocumentLimitBytes() {
    // LIM-021
    return REQUEST_DOCUMENT_LIMIT_BYTES;
  }

  /** Human-readable summary of the canonical JSON request document limit. */
  public static String requestDocumentLimitSummary() {
    // LIM-021
    return "request JSON must not exceed 16 MiB ("
        + REQUEST_DOCUMENT_LIMIT_BYTES
        + " bytes); use UTF8_FILE, FILE, or STANDARD_INPUT sources for large authored payloads.";
  }

  /** One stable product-owned message for oversized JSON request payloads. */
  public static String requestDocumentTooLargeMessage() {
    // LIM-021
    return "Request JSON exceeds the maximum size of 16 MiB ("
        + REQUEST_DOCUMENT_LIMIT_BYTES
        + " bytes); move large authored payloads into UTF8_FILE, FILE, or STANDARD_INPUT"
        + " sources.";
  }

  /** One stable step-kind explanation shared by help and discovery surfaces. */
  public static String stepKindSummary() {
    return "Every authored step requires a non-blank caller-defined stepId."
        + " stepId values must be unique within steps[] and must match [A-Za-z0-9._-]+."
        + " Use MUTATION steps for workbook changes, ASSERTION steps for first-class"
        + " verification, and INSPECTION steps for factual or analytical reads."
        + " Step kind is inferred from exactly one of action, assertion, or query;"
        + " do not send step.type.";
  }
}
