package dev.erst.gridgrind.excel;

/** Immutable snapshot of the protection currently applied to a cell style. */
public record ExcelCellProtectionSnapshot(boolean locked, boolean hiddenFormula) {}
