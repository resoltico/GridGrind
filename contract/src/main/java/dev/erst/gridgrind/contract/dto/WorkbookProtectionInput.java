package dev.erst.gridgrind.contract.dto;

/** Workbook-protection payload covering workbook and revisions lock state plus passwords. */
public record WorkbookProtectionInput(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionsLocked,
    String workbookPassword,
    String revisionsPassword) {
  public WorkbookProtectionInput {
    if (workbookPassword != null && workbookPassword.isBlank()) {
      throw new IllegalArgumentException("workbookPassword must not be blank");
    }
    if (revisionsPassword != null && revisionsPassword.isBlank()) {
      throw new IllegalArgumentException("revisionsPassword must not be blank");
    }
  }
}
