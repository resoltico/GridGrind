package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyValues;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.requireNonBlank;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
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
      StepId stepId, String sheetName, List<ExcelDrawingObjectSnapshot> drawingObjects)
      implements WorkbookDrawingResult {
    public DrawingObjectsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      drawingObjects = copyValues(drawingObjects, "drawingObjects");
    }
  }

  /** Returns factual chart metadata for one sheet. */
  record ChartsResult(StepId stepId, String sheetName, List<ChartReport> charts)
      implements WorkbookDrawingResult {
    public ChartsResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      charts = copyValues(charts, "charts");
    }
  }

  /** Returns factual pivot-table metadata selected by workbook-global pivot name or all pivots. */
  record PivotTablesResult(StepId stepId, List<PivotTableReport> pivotTables)
      implements WorkbookDrawingResult {
    public PivotTablesResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      pivotTables = copyValues(pivotTables, "pivotTables");
    }
  }

  /** Returns the extracted binary payload for one existing drawing object. */
  record DrawingObjectPayloadResult(
      StepId stepId, String sheetName, ExcelDrawingObjectPayload payload)
      implements WorkbookDrawingResult {
    public DrawingObjectPayloadResult {

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      sheetName = requireNonBlank(sheetName, "sheetName");
      Objects.requireNonNull(payload, "payload must not be null");
    }
  }
}
