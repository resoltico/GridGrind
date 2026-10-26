package dev.erst.gridgrind.core.model.protection;

import dev.erst.gridgrind.core.metadata.CatalogType;

/** Effective cell-protection facts reported with every analyzed cell. */
@CatalogType(
    id = "CellProtectionReport",
    summary = "Exact cell-protection report.",
    displayOrder = 82)
public record CellProtectionReport(boolean locked, boolean hiddenFormula) {}
