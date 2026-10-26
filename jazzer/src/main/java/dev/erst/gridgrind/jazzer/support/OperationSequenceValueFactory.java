package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.layout.PaneSpec;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.*;
import dev.erst.gridgrind.core.vocabulary.*;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.excel.*;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

/** Builds bounded protocol and engine value payloads shared across the Jazzer generators. */
final class OperationSequenceValueFactory {
  static final String DRAWING_PICTURE_NAME = "OpsPicture";
  static final String DRAWING_CHART_NAME = "OpsChart";
  static final String PIVOT_TABLE_NAME = "OpsPivot";
  static final String DRAWING_SHAPE_NAME = "OpsShape";
  static final String DRAWING_CONNECTOR_NAME = "OpsConnector";
  static final String DRAWING_EMBEDDED_OBJECT_NAME = "OpsEmbed";
  static final String PNG_PIXEL_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC";

  private OperationSequenceValueFactory() {}

  static PaneSpec nextPaneInput(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextPaneInput(data);
  }

  static ExcelSheetPane nextExcelSheetPane(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextExcelSheetPane(data);
  }

  static PrintLayoutSpec nextPrintLayoutInput(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextPrintLayoutInput(data);
  }

  static ResolvedPrintLayout nextExcelPrintLayout(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextExcelPrintLayout(data);
  }

  static ExcelPaneRegion nextProtocolPaneRegion(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextProtocolPaneRegion(data);
  }

  static ExcelPaneRegion nextPaneRegion(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextPaneRegion(data);
  }

  static SheetCopyPosition nextSheetCopyPosition(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextSheetCopyPosition(data);
  }

  static ExcelSheetCopyPosition nextExcelSheetCopyPosition(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextExcelSheetCopyPosition(data);
  }

  static List<String> nextSelectedSheetNames(
      GridGrindFuzzData data, String primarySheet, String secondarySheet) {
    return OperationSequenceLayoutValues.nextSelectedSheetNames(data, primarySheet, secondarySheet);
  }

  static ExcelSheetVisibility nextProtocolSheetVisibility(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextProtocolSheetVisibility(data);
  }

  static ExcelSheetVisibility nextSheetVisibility(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextSheetVisibility(data);
  }

  static SheetProtectionSettings nextProtocolSheetProtectionSettings(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextProtocolSheetProtectionSettings(data);
  }

  static SheetProtectionSettings nextSheetProtectionSettings(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextSheetProtectionSettings(data);
  }

  static OperationSequenceValueFactory.WorkflowStorage nextWorkflowStorage(
      String primarySheet, String secondarySheet, GridGrindFuzzData data) throws IOException {
    return OperationSequenceLayoutValues.nextWorkflowStorage(primarySheet, secondarySheet, data);
  }

  static void writeExistingWorkbook(
      Path sourcePath, String primarySheet, String secondarySheet, GridGrindFuzzData data)
      throws IOException {
    OperationSequenceLayoutValues.writeExistingWorkbook(
        sourcePath, primarySheet, secondarySheet, data);
  }

  static HyperlinkTarget nextHyperlinkTarget(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextHyperlinkTarget(data);
  }

  static HyperlinkTarget nextExcelHyperlink(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelHyperlink(data);
  }

  static CommentSpec nextCommentInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextCommentInput(data);
  }

  static PictureSpec nextPictureInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextPictureInput(data);
  }

  static ChartSpec nextChartInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextChartInput(data);
  }

  static ShapeSpec nextShapeInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextShapeInput(data);
  }

  static EmbeddedObjectSpec nextEmbeddedObjectInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextEmbeddedObjectInput(data);
  }

  static DrawingAnchorSpec.TwoCell nextDrawingAnchorInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextDrawingAnchorInput(data);
  }

  static PictureDataSpec nextPictureDataInput() {
    return OperationSequenceDrawingValues.nextPictureDataInput();
  }

  static ResolvedPicture nextExcelPictureDefinition(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelPictureDefinition(data);
  }

  static ResolvedChart nextResolvedChart(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextResolvedChart(data);
  }

  static ResolvedShape nextResolvedShape(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextResolvedShape(data);
  }

  static ResolvedEmbeddedObject nextExcelEmbeddedObjectDefinition(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelEmbeddedObjectDefinition(data);
  }

  static ExcelDrawingAnchorBehavior nextDrawingAnchorBehavior(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextDrawingAnchorBehavior(data);
  }

  static String nextDrawingObjectName(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextDrawingObjectName(data);
  }

  static String nextDrawingBinaryObjectName(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextDrawingBinaryObjectName(data);
  }

  static DataValidationSpec nextDataValidationInput(GridGrindFuzzData data) {
    return OperationSequenceRuleValues.nextDataValidationInput(data);
  }

  static ResolvedDataValidation nextExcelDataValidationDefinition(GridGrindFuzzData data) {
    return OperationSequenceRuleValues.nextExcelDataValidationDefinition(data);
  }

  static ConditionalFormattingSpec nextConditionalFormattingInput(GridGrindFuzzData data) {
    return OperationSequenceRuleValues.nextConditionalFormattingInput(data);
  }

  static WorkbookFormattingCommand.SetConditionalFormatting nextConditionalFormattingCommand(
      GridGrindFuzzData data, boolean validRange, String sheetName) {
    return OperationSequenceRuleValues.nextConditionalFormattingCommand(
        data, validRange, sheetName);
  }

  static DifferentialStylePatch nextDifferentialStyleInput(GridGrindFuzzData data) {
    return OperationSequenceRuleValues.nextDifferentialStyleInput(data);
  }

  static DifferentialStylePatch nextExcelDifferentialStyle(GridGrindFuzzData data) {
    return OperationSequenceRuleValues.nextExcelDifferentialStyle(data);
  }

  static RangeSelector nextRangeSelector(
      GridGrindFuzzData data, String sheetName, boolean validRange) {
    return OperationSequenceRuleValues.nextRangeSelector(data, sheetName, validRange);
  }

  static ExcelRangeSelection nextExcelRangeSelection(GridGrindFuzzData data, boolean validRange) {
    return OperationSequenceRuleValues.nextExcelRangeSelection(data, validRange);
  }

  static String nextAutofilterRange(boolean validRange) {
    return OperationSequenceRuleValues.nextAutofilterRange(validRange);
  }

  static String nextCopySheetName(String sourceSheetName) {
    return OperationSequenceRuleValues.nextCopySheetName(sourceSheetName);
  }

  static TableSpec nextTableInput(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return OperationSequenceTabularValues.nextTableInput(data, sheetName, tableName, validRange);
  }

  static ResolvedTable nextExcelTableDefinition(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return OperationSequenceTabularValues.nextExcelTableDefinition(
        data, sheetName, tableName, validRange);
  }

  static TableStyleSpec nextTableStyleInput(GridGrindFuzzData data) {
    return OperationSequenceTabularValues.nextTableStyleInput(data);
  }

  static TableStyleSpec nextExcelTableStyle(GridGrindFuzzData data) {
    return OperationSequenceTabularValues.nextExcelTableStyle(data);
  }

  static TableSelector nextTableSelector(
      GridGrindFuzzData data, String primarySheet, String secondarySheet) {
    return OperationSequenceTabularValues.nextTableSelector(data, primarySheet, secondarySheet);
  }

  static PivotTableSelector nextPivotTableSelector(
      GridGrindFuzzData data,
      String primarySheet,
      String secondarySheet,
      String pivotTableName,
      boolean validName) {
    return OperationSequenceTabularValues.nextPivotTableSelector(
        data, primarySheet, secondarySheet, pivotTableName, validName);
  }

  static PivotTableSpec nextPivotTableInput(
      GridGrindFuzzData data,
      String targetSheet,
      String pivotTableName,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return OperationSequenceTabularValues.nextPivotTableInput(
        data, targetSheet, pivotTableName, namedRangeName, tableName, validName, validRange);
  }

  static PivotTableSpec.PivotSourceSpec nextPivotTableSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return OperationSequenceTabularValues.nextPivotTableSource(
        data, targetSheet, namedRangeName, tableName, validName, validRange);
  }

  static PivotTableSpec nextPivotTableSpec(
      GridGrindFuzzData data,
      String targetSheet,
      String pivotTableName,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return OperationSequenceTabularValues.nextPivotTableSpec(
        data, targetSheet, pivotTableName, namedRangeName, tableName, validName, validRange);
  }

  static PivotTableSpec.PivotSourceSpec nextPivotSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return OperationSequenceTabularValues.nextPivotSource(
        data, targetSheet, namedRangeName, tableName, validName, validRange);
  }

  static String nextTableName(GridGrindFuzzData data, boolean valid, String sheetName) {
    return OperationSequenceTabularValues.nextTableName(data, valid, sheetName);
  }

  static ResolvedComment nextExcelComment(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelComment(data);
  }

  static String nextNamedRangeName(GridGrindFuzzData data, boolean valid) {
    return OperationSequenceTabularValues.nextNamedRangeName(data, valid);
  }

  static String nextPivotTableName(GridGrindFuzzData data, boolean valid) {
    return OperationSequenceTabularValues.nextPivotTableName(data, valid);
  }

  record WorkflowStorage(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      java.nio.file.Path cleanupRoot) {}
}
