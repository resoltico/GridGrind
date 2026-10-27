package dev.erst.gridgrind.engine.runtime.fixture;

/** Deliberately violates the native dependency boundary from runtime. */
public record ArchitectureNativeLibraryLeak(
    org.apache.poi.ss.usermodel.Workbook workbook,
    org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell cell) {}
