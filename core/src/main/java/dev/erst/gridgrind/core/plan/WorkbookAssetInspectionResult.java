package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
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
  @JsonTypeName("GET_DRAWING_OBJECTS")
  record DrawingObjectsResult(
      StepId stepId, String sheetName, List<DrawingObjectReport> drawingObjects)
      implements WorkbookAssetInspectionResult {
    public DrawingObjectsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      drawingObjects =
          InspectionResultValidationSupport.copyValues(drawingObjects, "drawingObjects");
    }
  }

  /** Returns factual chart metadata for one sheet. */
  @JsonTypeName("GET_CHARTS")
  record ChartsResult(StepId stepId, String sheetName, List<ChartReport> charts)
      implements WorkbookAssetInspectionResult {
    public ChartsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      charts = InspectionResultValidationSupport.copyValues(charts, "charts");
    }
  }

  /** Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots. */
  @JsonTypeName("GET_PIVOT_TABLES")
  record PivotTablesResult(StepId stepId, List<PivotTableReport> pivotTables)
      implements WorkbookAssetInspectionResult {
    public PivotTablesResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      pivotTables = InspectionResultValidationSupport.copyValues(pivotTables, "pivotTables");
    }
  }

  /** Returns the extracted binary payload for one existing drawing object. */
  @JsonTypeName("GET_DRAWING_OBJECT_PAYLOAD")
  record DrawingObjectPayloadResult(
      StepId stepId, String sheetName, DrawingObjectPayloadReport payload)
      implements WorkbookAssetInspectionResult {
    public DrawingObjectPayloadResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(payload, "payload must not be null");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  @JsonTypeName("GET_TABLES")
  record TablesResult(StepId stepId, List<TableEntryReport> tables)
      implements WorkbookAssetInspectionResult {
    public TablesResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      tables = InspectionResultValidationSupport.copyValues(tables, "tables");
    }
  }
}
