package dev.erst.gridgrind.excel.customxml;

/** Persisted custom-XML map behavior flags. */
public record ExcelCustomXmlMappingSettings(
    boolean showImportExportValidationErrors,
    boolean autoFit,
    boolean append,
    boolean preserveSortAfLayout,
    boolean preserveFormat) {}
