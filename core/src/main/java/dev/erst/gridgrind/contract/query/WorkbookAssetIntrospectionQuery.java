package dev.erst.gridgrind.contract.query;

import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;

/** Drawing-, chart-, pivot-, and table-scoped factual inspection queries. */
public sealed interface WorkbookAssetIntrospectionQuery extends InspectionQuery.Introspection
    permits WorkbookAssetIntrospectionQuery.GetDrawingObjects,
        WorkbookAssetIntrospectionQuery.GetCharts,
        WorkbookAssetIntrospectionQuery.GetPivotTables,
        WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload,
        WorkbookAssetIntrospectionQuery.GetTables {

  @ProtocolTypeMetadata(
      id = "GET_DRAWING_OBJECTS",
      summary = "Return factual drawing-object metadata for one sheet.",
      targetSelectors = {DrawingObjectSelector.AllOnSheet.class})
  record GetDrawingObjects() implements WorkbookAssetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_CHARTS",
      summary = "Return factual chart metadata for one sheet or one named chart.",
      targetSelectors = {ChartSelector.AllOnSheet.class, ChartSelector.ByName.class})
  record GetCharts() implements WorkbookAssetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_PIVOT_TABLES",
      summary = "Return factual pivot-table metadata selected by pivot-table name or ALL.",
      targetSelectors = {PivotTableSelector.class})
  record GetPivotTables() implements WorkbookAssetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_DRAWING_OBJECT_PAYLOAD",
      summary = "Return the extracted binary payload for one named picture or embedded object.",
      targetSelectors = {DrawingObjectSelector.ByName.class})
  record GetDrawingObjectPayload() implements WorkbookAssetIntrospectionQuery {}

  @ProtocolTypeMetadata(
      id = "GET_TABLES",
      summary = "Return factual table metadata selected by workbook-global table name or ALL.",
      targetSelectors = {TableSelector.class})
  record GetTables() implements WorkbookAssetIntrospectionQuery {}
}
