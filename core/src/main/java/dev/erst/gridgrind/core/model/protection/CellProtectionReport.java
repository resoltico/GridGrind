package dev.erst.gridgrind.core.model.protection;

/** Effective cell-protection facts reported with every analyzed cell. */
public record CellProtectionReport(boolean locked, boolean hiddenFormula) {}
