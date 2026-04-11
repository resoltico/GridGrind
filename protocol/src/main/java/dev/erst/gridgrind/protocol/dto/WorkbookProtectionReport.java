package dev.erst.gridgrind.protocol.dto;

/** Factual workbook-level protection state loaded from one workbook. */
public record WorkbookProtectionReport(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionLocked,
    boolean workbookPasswordHashPresent,
    boolean revisionsPasswordHashPresent) {}
