package dev.erst.gridgrind.core.fixture;

/** Deliberately violates both external-library boundaries of the canonical model. */
public record ArchitectureCoreLibraryLeak(
    org.apache.poi.ss.usermodel.Workbook workbook, tools.jackson.databind.JsonNode document) {}
