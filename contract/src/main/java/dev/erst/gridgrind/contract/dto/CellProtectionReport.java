package dev.erst.gridgrind.contract.dto;

/** Effective cell-protection facts reported with every analyzed cell. */
public record CellProtectionReport(boolean locked, boolean hiddenFormula) {}
