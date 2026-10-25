package dev.erst.gridgrind.core.model.protection;

/** Factual workbook-level protection state loaded from one workbook. */
public record WorkbookProtectionReport(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionsLocked,
    boolean workbookPasswordHashPresent,
    boolean revisionsPasswordHashPresent) {}
