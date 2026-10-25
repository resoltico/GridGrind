package dev.erst.gridgrind.contract.dto;

/** Factual workbook-level protection state loaded from one workbook. */
public record WorkbookProtectionReport(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionsLocked,
    boolean workbookPasswordHashPresent,
    boolean revisionsPasswordHashPresent) {}
