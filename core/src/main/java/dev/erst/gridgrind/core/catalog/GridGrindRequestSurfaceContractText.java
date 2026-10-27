package dev.erst.gridgrind.core.catalog;

/** Request size, workbook security and step semantics shared by public contract surfaces. */
public final class GridGrindRequestSurfaceContractText {
  private static final long REQUEST_DOCUMENT_LIMIT_BYTES = 16L * 1024 * 1024;

  private GridGrindRequestSurfaceContractText() {}

  /** Describes the host-independent execution-root rule for authored relative paths. */
  public static String requestOwnedPathResolutionSummary() {
    return "Relative request-owned paths resolve from the explicit execution root supplied by the host.";
  }

  /** Stable wording for encrypted OOXML plaintext temp ownership. */
  public static String encryptedWorkbookTempSecuritySummary() {
    return "Encrypted OOXML plaintext temp workbooks always stay in one private OS temporary"
        + " directory rather than the request root, execution root, or host scratch-parent"
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
