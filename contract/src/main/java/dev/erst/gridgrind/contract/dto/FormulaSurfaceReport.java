package dev.erst.gridgrind.contract.dto;

import java.util.List;

/** Grouped formula usage facts across one or more sheets. */
public record FormulaSurfaceReport(
    int totalFormulaCellCount, List<SheetFormulaSurfaceReport> sheets) {
  public FormulaSurfaceReport {
    sheets = GridGrindResponseSupport.copyValues(sheets, "sheets");
    if (totalFormulaCellCount < 0) {
      throw new IllegalArgumentException("totalFormulaCellCount must not be negative");
    }
  }
}
