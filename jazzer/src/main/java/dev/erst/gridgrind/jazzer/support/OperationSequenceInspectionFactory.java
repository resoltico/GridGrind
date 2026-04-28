package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.inspect;

import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;

/** Builds bounded inspection steps for protocol-workflow Jazzer generation. */
final class OperationSequenceInspectionFactory {
  private OperationSequenceInspectionFactory() {}

  static InspectionStep nextInspection(
      GridGrindFuzzData data,
      int index,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange,
      String pivotTableName) {
    String requestId = "read-" + index;
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    boolean validAddress = data.consumeBoolean();
    boolean validRange = data.consumeBoolean();
    boolean validName = data.consumeBoolean();
    int selector = nextSelectorByte(data);

    return switch (selectorFamily(selector)) {
      case 0x0 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary());
            case 0x1 ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetSheetSummary());
            case 0x2 ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextNamedRangeSelection(
                        data, targetSheet, workbookNamedRange, sheetNamedRange, validName),
                    new InspectionQuery.GetNamedRanges());
            default ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextNamedRangeSelection(
                        data, targetSheet, workbookNamedRange, sheetNamedRange, validName),
                    new InspectionQuery.GetNamedRangeSurface());
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    new CellSelector.ByAddresses(
                        targetSheet,
                        OperationSequenceObservationSupport.nextReadAddresses(data, validAddress)),
                    new InspectionQuery.GetCells());
            case 0x1 ->
                inspect(
                    requestId,
                    new RangeSelector.RectangularWindow(
                        targetSheet,
                        FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                        data.consumeInt(1, 4),
                        data.consumeInt(1, 4)),
                    new InspectionQuery.GetWindow());
            case 0x2 ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetSheetLayout());
            default ->
                inspect(
                    requestId,
                    new RangeSelector.RectangularWindow(
                        targetSheet,
                        validAddress ? "A1" : FuzzDataDecoders.nextNonBlankCellAddress(data, false),
                        data.consumeInt(1, 5),
                        data.consumeInt(1, 4)),
                    new InspectionQuery.GetSheetSchema());
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetMergedRegions());
            case 0x1 ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextCellSelector(
                        data, targetSheet, validAddress),
                    new InspectionQuery.GetHyperlinks());
            case 0x2 ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextCellSelector(
                        data, targetSheet, validAddress),
                    new InspectionQuery.GetComments());
            case 0x3 ->
                inspect(
                    requestId,
                    new DrawingObjectSelector.AllOnSheet(targetSheet),
                    new InspectionQuery.GetDrawingObjects());
            case 0x4 ->
                inspect(
                    requestId,
                    new DrawingObjectSelector.ByName(
                        targetSheet, nextDrawingBinaryObjectName(data)),
                    new InspectionQuery.GetDrawingObjectPayload());
            case 0x5 ->
                inspect(
                    requestId,
                    new ChartSelector.AllOnSheet(targetSheet),
                    new InspectionQuery.GetCharts());
            case 0x6 ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetPrintLayout());
            default ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextSheetSelector(
                        data, primarySheet, secondarySheet),
                    new InspectionQuery.GetFormulaSurface());
          };
      case 0x3 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    nextRangeSelector(data, targetSheet, validRange),
                    new InspectionQuery.GetDataValidations());
            case 0x1 ->
                inspect(
                    requestId,
                    nextRangeSelector(data, targetSheet, validRange),
                    new InspectionQuery.GetConditionalFormatting());
            default ->
                inspect(
                    requestId,
                    new SheetSelector.ByName(targetSheet),
                    new InspectionQuery.GetAutofilters());
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    nextTableSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.GetTables());
            case 0x1 ->
                inspect(
                    requestId,
                    nextPivotTableSelector(
                        data, primarySheet, secondarySheet, pivotTableName, validName),
                    new InspectionQuery.GetPivotTables());
            case 0x2 ->
                inspect(
                    requestId,
                    nextTableSelector(data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeTableHealth());
            case 0x3 ->
                inspect(
                    requestId,
                    nextPivotTableSelector(
                        data, primarySheet, secondarySheet, pivotTableName, validName),
                    new InspectionQuery.AnalyzePivotTableHealth());
            default ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextSheetSelector(
                        data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeAutofilterHealth());
          };
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextSheetSelector(
                        data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeFormulaHealth());
            case 0x1 ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextSheetSelector(
                        data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeDataValidationHealth());
            case 0x2 ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextSheetSelector(
                        data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeConditionalFormattingHealth());
            default ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextSheetSelector(
                        data, primarySheet, secondarySheet),
                    new InspectionQuery.AnalyzeHyperlinkHealth());
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                inspect(
                    requestId,
                    OperationSequenceObservationSupport.nextNamedRangeSelection(
                        data, targetSheet, workbookNamedRange, sheetNamedRange, validName),
                    new InspectionQuery.AnalyzeNamedRangeHealth());
            default ->
                inspect(
                    requestId,
                    new WorkbookSelector.Current(),
                    new InspectionQuery.AnalyzeWorkbookFindings());
          };
    };
  }
}
