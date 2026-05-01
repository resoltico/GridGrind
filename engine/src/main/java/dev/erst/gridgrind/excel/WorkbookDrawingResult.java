package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyValues;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.requireNonBlank;

import java.util.List;
import java.util.Objects;

/** Drawing and pivot inventory results. */
public sealed interface WorkbookDrawingResult extends WorkbookReadIntrospectionResult
    permits WorkbookDrawingResult.DrawingObjectsResult,
        WorkbookDrawingResult.ChartsResult,
        WorkbookDrawingResult.PivotTablesResult,
        WorkbookDrawingResult.DrawingObjectPayloadResult {

  /** Returns factual drawing-object metadata for one sheet. */
  record DrawingObjectsResult(
      String stepId, String sheetName, List<ExcelDrawingObjectSnapshot> drawingObjects)
      implements WorkbookDrawingResult {
    public DrawingObjectsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      drawingObjects = copyValues(drawingObjects, "drawingObjects");
    }
  }

  /** Returns factual chart metadata for one sheet. */
  record ChartsResult(String stepId, String sheetName, List<ExcelChartSnapshot> charts)
      implements WorkbookDrawingResult {
    public ChartsResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      charts = copyValues(charts, "charts");
    }
  }

  /** Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots. */
  record PivotTablesResult(String stepId, List<ExcelPivotTableSnapshot> pivotTables)
      implements WorkbookDrawingResult {
    public PivotTablesResult {
      stepId = requireNonBlank(stepId, "stepId");
      pivotTables = copyValues(pivotTables, "pivotTables");
    }
  }

  /** Returns the extracted binary payload for one existing drawing object. */
  record DrawingObjectPayloadResult(
      String stepId, String sheetName, ExcelDrawingObjectPayload payload)
      implements WorkbookDrawingResult {
    public DrawingObjectPayloadResult {
      stepId = requireNonBlank(stepId, "stepId");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(payload, "payload must not be null");
    }
  }
}
