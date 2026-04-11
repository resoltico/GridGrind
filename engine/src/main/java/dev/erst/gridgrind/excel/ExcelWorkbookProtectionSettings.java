package dev.erst.gridgrind.excel;

/** Mutable-workbook protection payload covering workbook and revisions locks plus passwords. */
public record ExcelWorkbookProtectionSettings(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionsLocked,
    String workbookPassword,
    String revisionsPassword) {
  public ExcelWorkbookProtectionSettings {
    if (workbookPassword != null && workbookPassword.isBlank()) {
      throw new IllegalArgumentException("workbookPassword must not be blank");
    }
    if (revisionsPassword != null && revisionsPassword.isBlank()) {
      throw new IllegalArgumentException("revisionsPassword must not be blank");
    }
  }
}
