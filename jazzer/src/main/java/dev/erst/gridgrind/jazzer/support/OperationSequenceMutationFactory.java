package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.mutate;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.List;

/** Builds bounded mutation steps for protocol-workflow Jazzer generation. */
final class OperationSequenceMutationFactory {
  private OperationSequenceMutationFactory() {}

  static ProtocolStepSupport.PendingMutation nextMutation(
      GridGrindFuzzData data,
      String primarySheet,
      String secondarySheet,
      String workbookNamedRange,
      String sheetNamedRange,
      String pivotTableName) {
    String targetSheet = data.consumeBoolean() ? primarySheet : secondarySheet;
    boolean validAddress = data.consumeBoolean();
    boolean validRange = data.consumeBoolean();
    boolean validName = data.consumeBoolean();
    IndexSpan columnSpan = nextIndexSpan(data, 3);
    IndexSpan rowSpan = nextIndexSpan(data, 3);
    String namedRangeName = data.consumeBoolean() ? workbookNamedRange : sheetNamedRange;
    String tableName = nextTableName(data, validName, targetSheet);
    int selector = nextSelectorByte(data);

    return switch (selectorFamily(selector)) {
      case 0x0 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.EnsureSheet());
            case 0x1 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.RenameSheet(primarySheet + "Renamed"));
            case 0x2 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.DeleteSheet());
            case 0x3 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.MoveSheet(data.consumeInt(0, 2)));
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.CopySheet(
                        nextCopySheetName(targetSheet), nextSheetCopyPosition(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.SetActiveSheet());
            case 0x6 ->
                mutate(
                    new SheetSelector.ByNames(
                        nextSelectedSheetNames(data, primarySheet, secondarySheet)),
                    new WorkbookMutationAction.SetSelectedSheets());
            case 0x7 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.SetSheetVisibility(
                        nextProtocolSheetVisibility(data)));
            case 0x8 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.SetSheetProtection(
                        nextProtocolSheetProtectionSettings(data)));
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.ClearSheetProtection());
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange)),
                    new WorkbookMutationAction.MergeCells());
            case 0x1 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange)),
                    new WorkbookMutationAction.UnmergeCells());
            case 0x2 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.SetColumnWidth(
                        data.consumeRegularDouble(1.0d, 20.0d)));
            case 0x3 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.SetRowHeight(
                        data.consumeRegularDouble(5.0d, 40.0d)));
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.SetSheetPane(nextPaneInput(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.SetSheetZoom(data.consumeInt(10, 400)));
            case 0x6 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.SetPrintLayout(nextPrintLayoutInput(data)));
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.ClearPrintLayout());
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new CellMutationAction.SetCell(FuzzDataDecoders.nextCellInput(data)));
            case 0x1 -> {
              String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
              yield mutate(
                  new RangeSelector.ByRange(targetSheet, range),
                  new CellMutationAction.SetRange(FuzzDataDecoders.nextProtocolMatrix(data, 2, 2)));
            }
            case 0x2 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange)),
                    new CellMutationAction.ClearRange());
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new CellMutationAction.AppendRow(
                        List.of(
                            FuzzDataDecoders.nextCellInput(data),
                            FuzzDataDecoders.nextCellInput(data))));
          };
      case 0x3 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new CellMutationAction.SetHyperlink(nextHyperlinkTarget(data)));
            case 0x1 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new CellMutationAction.ClearHyperlink());
            case 0x2 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new CellMutationAction.SetComment(nextCommentInput(data)));
            case 0x3 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new CellMutationAction.ClearComment());
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new DrawingMutationAction.SetPicture(nextPictureInput(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new DrawingMutationAction.SetChart(nextChartInput(data)));
            case 0x6 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new DrawingMutationAction.SetShape(nextShapeInput(data)));
            case 0x7 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new DrawingMutationAction.SetEmbeddedObject(nextEmbeddedObjectInput(data)));
            case 0x8 ->
                mutate(
                    new DrawingObjectSelector.ByName(targetSheet, nextDrawingObjectName(data)),
                    new DrawingMutationAction.SetDrawingObjectAnchor(nextDrawingAnchorInput(data)));
            case 0x9 ->
                mutate(
                    new DrawingObjectSelector.ByName(targetSheet, nextDrawingObjectName(data)),
                    new DrawingMutationAction.DeleteDrawingObject());
            default ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet,
                        validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false)),
                    new CellMutationAction.ApplyStyle(WorkbookStyleInputs.nextStyleInput(data)));
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet,
                        validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
                    new StructuredMutationAction.SetDataValidation(nextDataValidationInput(data)));
            case 0x1 ->
                mutate(
                    nextRangeSelector(data, targetSheet, validRange),
                    new StructuredMutationAction.ClearDataValidations());
            case 0x2 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new StructuredMutationAction.SetConditionalFormatting(
                        nextConditionalFormattingInput(data, validRange)));
            default ->
                mutate(
                    nextRangeSelector(data, targetSheet, validRange),
                    new StructuredMutationAction.ClearConditionalFormatting());
          };
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(targetSheet, nextAutofilterRange(validRange)),
                    new StructuredMutationAction.SetAutofilter());
            case 0x1 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new StructuredMutationAction.ClearAutofilter());
            case 0x2 ->
                mutate(
                    new StructuredMutationAction.SetTable(
                        nextTableInput(data, targetSheet, tableName, validRange)));
            case 0x3 ->
                mutate(
                    new TableSelector.ByNameOnSheet(tableName, targetSheet),
                    new StructuredMutationAction.DeleteTable());
            case 0x4 ->
                mutate(
                    new StructuredMutationAction.SetPivotTable(
                        nextPivotTableInput(
                            data,
                            targetSheet,
                            pivotTableName,
                            workbookNamedRange,
                            tableName,
                            validName,
                            validRange)));
            default ->
                mutate(
                    new PivotTableSelector.ByNameOnSheet(
                        validName ? pivotTableName : nextPivotTableName(data, false), targetSheet),
                    new StructuredMutationAction.DeletePivotTable());
          };
      case 0x6 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new StructuredMutationAction.SetNamedRange(
                        validName ? namedRangeName : nextNamedRangeName(data, false),
                        data.consumeBoolean()
                            ? new NamedRangeScope.Workbook()
                            : new NamedRangeScope.Sheet(targetSheet),
                        new NamedRangeTarget(
                            targetSheet,
                            validRange
                                ? "A1:B2"
                                : FuzzDataDecoders.nextNonBlankRange(data, false))));
            default ->
                mutate(
                    data.consumeBoolean()
                        ? new NamedRangeSelector.WorkbookScope(
                            validName ? namedRangeName : nextNamedRangeName(data, false))
                        : new NamedRangeSelector.SheetScope(
                            validName ? namedRangeName : nextNamedRangeName(data, false),
                            targetSheet),
                    new StructuredMutationAction.DeleteNamedRange());
          };
      case 0x7 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RowBandSelector.Insertion(
                        targetSheet, rowSpan.first(), rowSpan.last() - rowSpan.first() + 1),
                    new WorkbookMutationAction.InsertRows());
            case 0x1 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.DeleteRows());
            case 0x2 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.ShiftRows(nextNonZeroDelta(data, 2)));
            case 0x3 ->
                mutate(
                    new ColumnBandSelector.Insertion(
                        targetSheet,
                        columnSpan.first(),
                        columnSpan.last() - columnSpan.first() + 1),
                    new WorkbookMutationAction.InsertColumns());
            case 0x4 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.DeleteColumns());
            case 0x5 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.ShiftColumns(nextNonZeroDelta(data, 2)));
            case 0x6 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.SetRowVisibility(data.consumeBoolean()));
            case 0x7 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.SetColumnVisibility(data.consumeBoolean()));
            case 0x8 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.GroupRows(data.consumeBoolean()));
            case 0x9 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.UngroupRows());
            case 0xA ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.GroupColumns(data.consumeBoolean()));
            default ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.UngroupColumns());
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.AutoSizeColumns());
            case 0x1 ->
                mutate(
                    new WorkbookSelector.Current(),
                    new WorkbookMutationAction.SetWorkbookProtection(
                        new WorkbookProtectionInput(true, false, false, null, null)));
            case 0x2 ->
                mutate(
                    new WorkbookSelector.Current(),
                    new WorkbookMutationAction.ClearWorkbookProtection());
            case 0x3 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.SetSheetPresentation(
                        SheetPresentationInput.defaults()));
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new WorkbookMutationAction.AutoSizeColumns());
          };
    };
  }
}
