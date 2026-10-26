package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableSpec;

/** Converts table and pivot-table contract inputs. */
final class WorkbookCommandTabularInputConverter {
  private WorkbookCommandTabularInputConverter() {}

  static ResolvedTable toResolvedTable(TableSpec table) {
    return new ResolvedTable(
        table.name(),
        table.sheetName(),
        table.range(),
        table.showTotalsRow(),
        table.hasAutofilter(),
        table.style(),
        WorkbookCommandSourceSupport.inlineText(table.comment(), "table comment"),
        table.published(),
        table.insertRow(),
        table.insertRowShift(),
        table.headerRowCellStyle(),
        table.dataCellStyle(),
        table.totalsRowCellStyle(),
        table.columns());
  }
}
