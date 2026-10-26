package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import java.util.Objects;

/** Workbook asset and object inspection results. */
public sealed interface WorkbookAssetInspectionResult extends InspectionIntrospectionResult
    permits WorkbookAssetInspectionResult.DrawingObjectsResult,
        WorkbookAssetInspectionResult.ChartsResult,
        WorkbookAssetInspectionResult.PivotTablesResult,
        WorkbookAssetInspectionResult.DrawingObjectPayloadResult,
        WorkbookAssetInspectionResult.TablesResult {

  /** Returns factual drawing-object metadata for one sheet. */
  @Variant(
      id = "GET_DRAWING_OBJECTS",
      summary = "Returns factual drawing-object metadata for one sheet.")
  record DrawingObjectsResult(
      StepId stepId, SheetLabel sheetName, List<DrawingObjectReport> drawingObjects)
      implements WorkbookAssetInspectionResult {
    public DrawingObjectsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      drawingObjects =
          InspectionResultValidationSupport.copyValues(drawingObjects, "drawingObjects");
    }
  }

  /** Returns factual chart metadata for one sheet. */
  @Variant(id = "GET_CHARTS", summary = "Returns factual chart metadata for one sheet.")
  record ChartsResult(StepId stepId, SheetLabel sheetName, List<ChartReport> charts)
      implements WorkbookAssetInspectionResult {
    public ChartsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      charts = InspectionResultValidationSupport.copyValues(charts, "charts");
    }
  }

  /** Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots. */
  @Variant(
      id = "GET_PIVOT_TABLES",
      summary =
          "Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots.")
  record PivotTablesResult(StepId stepId, List<PivotTableReport> pivotTables)
      implements WorkbookAssetInspectionResult {
    public PivotTablesResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      pivotTables = InspectionResultValidationSupport.copyValues(pivotTables, "pivotTables");
    }
  }

  /** Returns the extracted binary payload for one existing drawing object. */
  @Variant(
      id = "GET_DRAWING_OBJECT_PAYLOAD",
      summary = "Returns the extracted binary payload for one existing drawing object.")
  record DrawingObjectPayloadResult(
      StepId stepId, SheetLabel sheetName, DrawingObjectPayloadReport payload)
      implements WorkbookAssetInspectionResult {
    public DrawingObjectPayloadResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(payload, "payload must not be null");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  @Variant(
      id = "GET_TABLES",
      summary =
          "Returns factual table metadata selected by workbook-global table name or all tables.")
  record TablesResult(StepId stepId, List<TableEntryReport> tables)
      implements WorkbookAssetInspectionResult {
    public TablesResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      tables = InspectionResultValidationSupport.copyValues(tables, "tables");
    }
  }
}
