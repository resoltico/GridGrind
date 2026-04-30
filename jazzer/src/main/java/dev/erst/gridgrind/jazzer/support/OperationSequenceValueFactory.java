package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.source.*;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.foundation.*;
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
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  private OperationSequenceValueFactory() {}

  static PaneInput nextPaneInput(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextPaneInput(data);
  }

  static ExcelSheetPane nextExcelSheetPane(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextExcelSheetPane(data);
  }

  static PrintLayoutInput nextPrintLayoutInput(GridGrindFuzzData data) {
    return OperationSequenceLayoutValues.nextPrintLayoutInput(data);
  }

  static ExcelPrintLayout nextExcelPrintLayout(GridGrindFuzzData data) {
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

  static ExcelSheetProtectionSettings nextSheetProtectionSettings(GridGrindFuzzData data) {
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

  static ExcelHyperlink nextExcelHyperlink(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelHyperlink(data);
  }

  static CommentInput nextCommentInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextCommentInput(data);
  }

  static PictureInput nextPictureInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextPictureInput(data);
  }

  static ChartInput nextChartInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextChartInput(data);
  }

  static ShapeInput nextShapeInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextShapeInput(data);
  }

  static EmbeddedObjectInput nextEmbeddedObjectInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextEmbeddedObjectInput(data);
  }

  static DrawingAnchorInput.TwoCell nextDrawingAnchorInput(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextDrawingAnchorInput(data);
  }

  static PictureDataInput nextPictureDataInput() {
    return OperationSequenceDrawingValues.nextPictureDataInput();
  }

  static ExcelPictureDefinition nextExcelPictureDefinition(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelPictureDefinition(data);
  }

  static ExcelChartDefinition nextExcelChartDefinition(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelChartDefinition(data);
  }

  static ExcelShapeDefinition nextExcelShapeDefinition(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelShapeDefinition(data);
  }

  static ExcelEmbeddedObjectDefinition nextExcelEmbeddedObjectDefinition(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelEmbeddedObjectDefinition(data);
  }

  static ExcelDrawingAnchor.TwoCell nextExcelDrawingAnchor(GridGrindFuzzData data) {
    return OperationSequenceDrawingValues.nextExcelDrawingAnchor(data);
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

  static DataValidationInput nextDataValidationInput(GridGrindFuzzData data) {
    return OperationSequenceRuleValues.nextDataValidationInput(data);
  }

  static ExcelDataValidationDefinition nextExcelDataValidationDefinition(GridGrindFuzzData data) {
    return OperationSequenceRuleValues.nextExcelDataValidationDefinition(data);
  }

  static ConditionalFormattingBlockInput nextConditionalFormattingInput(
      GridGrindFuzzData data, boolean validRange) {
    return OperationSequenceRuleValues.nextConditionalFormattingInput(data, validRange);
  }

  static ExcelConditionalFormattingBlockDefinition nextExcelConditionalFormattingBlockDefinition(
      GridGrindFuzzData data, boolean validRange) {
    return OperationSequenceRuleValues.nextExcelConditionalFormattingBlockDefinition(
        data, validRange);
  }

  static DifferentialStyleInput nextDifferentialStyleInput(GridGrindFuzzData data) {
    return OperationSequenceRuleValues.nextDifferentialStyleInput(data);
  }

  static ExcelDifferentialStyle nextExcelDifferentialStyle(GridGrindFuzzData data) {
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

  static TableInput nextTableInput(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return OperationSequenceTabularValues.nextTableInput(data, sheetName, tableName, validRange);
  }

  static ExcelTableDefinition nextExcelTableDefinition(
      GridGrindFuzzData data, String sheetName, String tableName, boolean validRange) {
    return OperationSequenceTabularValues.nextExcelTableDefinition(
        data, sheetName, tableName, validRange);
  }

  static TableStyleInput nextTableStyleInput(GridGrindFuzzData data) {
    return OperationSequenceTabularValues.nextTableStyleInput(data);
  }

  static ExcelTableStyle nextExcelTableStyle(GridGrindFuzzData data) {
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

  static PivotTableInput nextPivotTableInput(
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

  static PivotTableInput.Source nextPivotTableSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return OperationSequenceTabularValues.nextPivotTableSource(
        data, targetSheet, namedRangeName, tableName, validName, validRange);
  }

  static ExcelPivotTableDefinition nextExcelPivotTableDefinition(
      GridGrindFuzzData data,
      String targetSheet,
      String pivotTableName,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return OperationSequenceTabularValues.nextExcelPivotTableDefinition(
        data, targetSheet, pivotTableName, namedRangeName, tableName, validName, validRange);
  }

  static ExcelPivotTableDefinition.Source nextExcelPivotTableSource(
      GridGrindFuzzData data,
      String targetSheet,
      String namedRangeName,
      String tableName,
      boolean validName,
      boolean validRange) {
    return OperationSequenceTabularValues.nextExcelPivotTableSource(
        data, targetSheet, namedRangeName, tableName, validName, validRange);
  }

  static String nextTableName(GridGrindFuzzData data, boolean valid, String sheetName) {
    return OperationSequenceTabularValues.nextTableName(data, valid, sheetName);
  }

  static ExcelComment nextExcelComment(GridGrindFuzzData data) {
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
