package dev.erst.gridgrind.core.model.protection;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Factual workbook-level protection state loaded from one workbook. */
@CatalogType(
    id = "WorkbookProtectionReport",
    summary =
        "Exact workbook-protection report covering structure, windows, revisions, and password-hash presence flags.",
    displayOrder = 76)
public record WorkbookProtectionReport(
    boolean structureLocked,
    boolean windowsLocked,
    boolean revisionsLocked,
    boolean workbookPasswordHashPresent,
    boolean revisionsPasswordHashPresent) {}
