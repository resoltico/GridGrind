package dev.erst.gridgrind.excel;

/** Immutable workbook-level protection facts loaded from one workbook. */
public record ExcelWorkbookProtectionSnapshot(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionLocked,
    boolean workbookPasswordHashPresent,
    boolean revisionsPasswordHashPresent) {}
