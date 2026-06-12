package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;

/** Canonical workbook-asset query factories kept internal to the Java authoring surface. */
final class WorkbookAssetQueries {
  private WorkbookAssetQueries() {}

  static WorkbookAssetIntrospectionQuery.GetDrawingObjects drawingObjects() {
    return new WorkbookAssetIntrospectionQuery.GetDrawingObjects();
  }

  static WorkbookAssetIntrospectionQuery.GetCharts charts() {
    return new WorkbookAssetIntrospectionQuery.GetCharts();
  }

  static WorkbookAssetIntrospectionQuery.GetPivotTables pivotTables() {
    return new WorkbookAssetIntrospectionQuery.GetPivotTables();
  }

  static WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload drawingObjectPayload() {
    return new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload();
  }

  static WorkbookAssetIntrospectionQuery.GetTables tables() {
    return new WorkbookAssetIntrospectionQuery.GetTables();
  }
}
