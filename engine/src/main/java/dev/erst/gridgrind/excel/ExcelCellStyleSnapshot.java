package dev.erst.gridgrind.excel;

/** Immutable snapshot of the style currently applied to a cell. */
public record ExcelCellStyleSnapshot(
    String numberFormat,
    boolean bold,
    boolean italic,
    boolean wrapText,
    String horizontalAlignment,
    String verticalAlignment) {}
