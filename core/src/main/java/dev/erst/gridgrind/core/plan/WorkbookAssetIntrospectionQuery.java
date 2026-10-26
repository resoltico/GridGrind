package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;

/** Drawing-, chart-, pivot-, and table-scoped factual inspection queries. */
public sealed interface WorkbookAssetIntrospectionQuery extends InspectionQuery.Introspection
    permits WorkbookAssetIntrospectionQuery.GetDrawingObjects,
        WorkbookAssetIntrospectionQuery.GetCharts,
        WorkbookAssetIntrospectionQuery.GetPivotTables,
        WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload,
        WorkbookAssetIntrospectionQuery.GetTables {

  @Variant(
      id = "GET_DRAWING_OBJECTS",
      summary = "Return factual drawing-object metadata for one sheet.",
      targetSelectors = {DrawingObjectSelector.AllOnSheet.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetDrawingObjects() implements WorkbookAssetIntrospectionQuery {}

  @Variant(
      id = "GET_CHARTS",
      summary = "Return factual chart metadata for one sheet or one named chart.",
      targetSelectors = {ChartSelector.AllOnSheet.class, ChartSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetCharts() implements WorkbookAssetIntrospectionQuery {}

  @Variant(
      id = "GET_PIVOT_TABLES",
      summary = "Return factual pivot-table metadata selected by pivot-table name or ALL.",
      targetSelectors = {PivotTableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetPivotTables() implements WorkbookAssetIntrospectionQuery {}

  @Variant(
      id = "GET_DRAWING_OBJECT_PAYLOAD",
      summary = "Return the extracted binary payload for one named picture or embedded object.",
      targetSelectors = {DrawingObjectSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetDrawingObjectPayload() implements WorkbookAssetIntrospectionQuery {}

  @Variant(
      id = "GET_TABLES",
      summary = "Return factual table metadata selected by workbook-global table name or ALL.",
      targetSelectors = {TableSelector.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.READ)
  record GetTables() implements WorkbookAssetIntrospectionQuery {}
}
