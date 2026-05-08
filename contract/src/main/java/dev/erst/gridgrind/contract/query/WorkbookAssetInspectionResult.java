package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectPayloadReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
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
      String stepId, String sheetName, List<DrawingObjectReport> drawingObjects)
      implements WorkbookAssetInspectionResult {
    public DrawingObjectsResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      drawingObjects =
          InspectionResultValidationSupport.copyValues(drawingObjects, "drawingObjects");
    }
  }

  /** Returns factual chart metadata for one sheet. */
  @JsonTypeName("GET_CHARTS")
  record ChartsResult(String stepId, String sheetName, List<ChartReport> charts)
      implements WorkbookAssetInspectionResult {
    public ChartsResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      charts = InspectionResultValidationSupport.copyValues(charts, "charts");
    }
  }

  /** Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots. */
  @JsonTypeName("GET_PIVOT_TABLES")
  record PivotTablesResult(String stepId, List<PivotTableReport> pivotTables)
      implements WorkbookAssetInspectionResult {
    public PivotTablesResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      pivotTables = InspectionResultValidationSupport.copyValues(pivotTables, "pivotTables");
    }
  }

  /** Returns the extracted binary payload for one existing drawing object. */
  @JsonTypeName("GET_DRAWING_OBJECT_PAYLOAD")
  record DrawingObjectPayloadResult(
      String stepId, String sheetName, DrawingObjectPayloadReport payload)
      implements WorkbookAssetInspectionResult {
    public DrawingObjectPayloadResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      sheetName = InspectionResultValidationSupport.requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(payload, "payload must not be null");
    }
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  @JsonTypeName("GET_TABLES")
  record TablesResult(String stepId, List<TableEntryReport> tables)
      implements WorkbookAssetInspectionResult {
    public TablesResult {
      stepId = InspectionResultValidationSupport.requireNonBlank(stepId, "stepId");
      tables = InspectionResultValidationSupport.copyValues(tables, "tables");
    }
  }
}
