package dev.erst.gridgrind.contract.dto;

/** Workbook-protection payload covering workbook and revisions lock state plus passwords. */
public record WorkbookProtectionInput(
    Boolean structureLocked,
    Boolean windowsLocked,
    Boolean revisionsLocked,
    String workbookPassword,
    String revisionsPassword) {
  public WorkbookProtectionInput {
    structureLocked = Boolean.TRUE.equals(structureLocked);
    windowsLocked = Boolean.TRUE.equals(windowsLocked);
    revisionsLocked = Boolean.TRUE.equals(revisionsLocked);
    if (workbookPassword != null && workbookPassword.isBlank()) {
      throw new IllegalArgumentException("workbookPassword must not be blank");
    }
    if (revisionsPassword != null && revisionsPassword.isBlank()) {
      throw new IllegalArgumentException("revisionsPassword must not be blank");
    }
  }
}
