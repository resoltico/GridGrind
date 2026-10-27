package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.result.WorkbookResultSupport;
import java.util.List;

/** Grouped formula usage facts across one or more sheets. */
public record FormulaSurfaceReport(
    int totalFormulaCellCount, List<SheetFormulaSurfaceReport> sheets) {
  public FormulaSurfaceReport {
    sheets = WorkbookResultSupport.copyValues(sheets, "sheets");
    if (totalFormulaCellCount < 0) {
      throw new IllegalArgumentException("totalFormulaCellCount must not be negative");
    }
  }
}
