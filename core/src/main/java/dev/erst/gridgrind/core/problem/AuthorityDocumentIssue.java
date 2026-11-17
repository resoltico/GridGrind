package dev.erst.gridgrind.core.problem;

/** Safe authority-intake wording, distinct from a valid grant denying an action. */
public enum AuthorityDocumentIssue {
  SYNTAX("Authority document is not valid JSON."),
  SHAPE("Authority document does not match the declared grant structure."),
  VALUE("Authority document contains a value that violates a declared grant constraint.");

  private final String message;

  AuthorityDocumentIssue(String message) {
    this.message = message;
  }

  /** Returns safe wording without echoing document values or provider exception text. */
  public String message() {
    return message;
  }
}
