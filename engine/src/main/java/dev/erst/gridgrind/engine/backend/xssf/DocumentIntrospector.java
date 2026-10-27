package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.engine.autofilter.AutofilterAdapter;
import dev.erst.gridgrind.engine.pivots.PivotAdapter;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.util.List;
import java.util.Objects;

/** Reads document-intelligence facts from one sheet wrapper. */
final class DocumentIntrospector {
  private final AutofilterAdapter autofilterAdapter;
  private final TableAdapter tableController;
  private final PivotAdapter pivotTableController;

  DocumentIntrospector() {
    this(new AutofilterAdapter(), new TableAdapter(), new PivotAdapter());
  }

  DocumentIntrospector(
      AutofilterAdapter autofilterAdapter,
      TableAdapter tableController,
      PivotAdapter pivotTableController) {
    this.autofilterAdapter =
        Objects.requireNonNull(autofilterAdapter, "autofilterAdapter must not be null");
    this.tableController =
        Objects.requireNonNull(tableController, "tableController must not be null");
    this.pivotTableController =
        Objects.requireNonNull(pivotTableController, "pivotTableController must not be null");
  }

  /** Returns sheet- and table-owned autofilter metadata for one sheet. */
  List<AutofilterEntryReport> autofilters(XssfWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");

    XssfSheet sheet = workbook.sheet(sheetName);
    List<AutofilterEntryReport> autofilters = new java.util.ArrayList<>();
    autofilters.addAll(autofilterAdapter.sheetOwnedAutofilters(sheet.xssfSheet()));
    autofilters.addAll(tableController.tableOwnedAutofilters(workbook, sheetName));
    return List.copyOf(autofilters);
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  List<TableEntryReport> tables(XssfWorkbook workbook, TableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return tableController.tables(workbook, selection);
  }

  /** Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots. */
  List<PivotTableReport> pivotTables(XssfWorkbook workbook, PivotTableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return pivotTableController.pivotTables(workbook, selection);
  }

  /** Returns factual drawing-object metadata for one sheet. */
  List<DrawingObjectReport> drawingObjects(XssfWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    return workbook.sheet(sheetName).drawings().drawingObjects();
  }

  /** Returns factual chart metadata selected by sheet-wide or exact chart selection. */
  List<ChartReport> charts(XssfWorkbook workbook, ChartSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    List<ChartReport> charts = workbook.sheet(selection.sheetName().value()).drawings().charts();
    return switch (selection) {
      case ChartSelector.AllOnSheet _ -> charts;
      case ChartSelector.ByName byName ->
          charts.stream().filter(chart -> chart.name().equals(byName.chartName())).toList();
    };
  }

  /** Returns the extracted binary payload for one existing drawing object on one sheet. */
  DrawingObjectPayloadReport drawingObjectPayload(
      XssfWorkbook workbook, String sheetName, String objectName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(objectName, "objectName must not be null");
    return workbook.sheet(sheetName).drawings().drawingObjectPayload(objectName);
  }
}
