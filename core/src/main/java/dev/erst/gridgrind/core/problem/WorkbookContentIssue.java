package dev.erst.gridgrind.core.problem;

/** Closed, safe reasons for native workbook-content refusals. */
public enum WorkbookContentIssue {
  XML_SOURCE_REFUSED(
      GridGrindProblemCode.INVALID_WORKBOOK_CONTENT,
      "Workbook XML is malformed or contains declarations this reader does not accept."),
  INVALID_NATIVE_VALUE(
      GridGrindProblemCode.INVALID_WORKBOOK_CONTENT,
      "Workbook content contains a value outside its native schema."),
  UNSUPPORTED_PRESERVATION(
      GridGrindProblemCode.UNSUPPORTED_WORKBOOK_CONTENT,
      "The requested operation cannot preserve this native workbook structure."),
  UNSUPPORTED_DEPENDENCY(
      GridGrindProblemCode.UNSUPPORTED_WORKBOOK_CONTENT,
      "The requested operation cannot preserve a required native relationship.");

  private final GridGrindProblemCode code;
  private final String message;

  WorkbookContentIssue(GridGrindProblemCode code, String message) {
    this.code = code;
    this.message = message;
  }

  /** Returns the core-owned refusal kind. */
  public GridGrindProblemCode problemCode() {
    return code;
  }

  /** Returns safe wording without native payload or exception text. */
  public String message() {
    return message;
  }
}
