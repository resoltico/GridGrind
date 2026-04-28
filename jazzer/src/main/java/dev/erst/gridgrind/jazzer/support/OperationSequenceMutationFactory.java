package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.mutate;

import dev.erst.gridgrind.contract.action.MutationAction;
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
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.EnsureSheet());
            case 0x1 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.RenameSheet(primarySheet + "Renamed"));
            case 0x2 ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.DeleteSheet());
            case 0x3 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.MoveSheet(data.consumeInt(0, 2)));
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.CopySheet(
                        nextCopySheetName(targetSheet), nextSheetCopyPosition(data)));
            case 0x5 ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.SetActiveSheet());
            case 0x6 ->
                mutate(
                    new SheetSelector.ByNames(
                        nextSelectedSheetNames(data, primarySheet, secondarySheet)),
                    new MutationAction.SetSelectedSheets());
            case 0x7 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetVisibility(nextProtocolSheetVisibility(data)));
            case 0x8 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetProtection(
                        nextProtocolSheetProtectionSettings(data)));
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.ClearSheetProtection());
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange)),
                    new MutationAction.MergeCells());
            case 0x1 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange)),
                    new MutationAction.UnmergeCells());
            case 0x2 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.SetColumnWidth(data.consumeRegularDouble(1.0d, 20.0d)));
            case 0x3 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.SetRowHeight(data.consumeRegularDouble(5.0d, 40.0d)));
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetPane(nextPaneInput(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetZoom(data.consumeInt(10, 400)));
            case 0x6 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetPrintLayout(nextPrintLayoutInput(data)));
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet), new MutationAction.ClearPrintLayout());
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new MutationAction.SetCell(FuzzDataDecoders.nextCellInput(data)));
            case 0x1 -> {
              String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
              yield mutate(
                  new RangeSelector.ByRange(targetSheet, range),
                  new MutationAction.SetRange(FuzzDataDecoders.nextProtocolMatrix(data, 2, 2)));
            }
            case 0x2 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange)),
                    new MutationAction.ClearRange());
            default ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.AppendRow(
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
                    new MutationAction.SetHyperlink(nextHyperlinkTarget(data)));
            case 0x1 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new MutationAction.ClearHyperlink());
            case 0x2 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new MutationAction.SetComment(nextCommentInput(data)));
            case 0x3 ->
                mutate(
                    new CellSelector.ByAddress(
                        targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress)),
                    new MutationAction.ClearComment());
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetPicture(nextPictureInput(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetChart(nextChartInput(data)));
            case 0x6 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetShape(nextShapeInput(data)));
            case 0x7 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetEmbeddedObject(nextEmbeddedObjectInput(data)));
            case 0x8 ->
                mutate(
                    new DrawingObjectSelector.ByName(targetSheet, nextDrawingObjectName(data)),
                    new MutationAction.SetDrawingObjectAnchor(nextDrawingAnchorInput(data)));
            case 0x9 ->
                mutate(
                    new DrawingObjectSelector.ByName(targetSheet, nextDrawingObjectName(data)),
                    new MutationAction.DeleteDrawingObject());
            default ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet,
                        validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false)),
                    new MutationAction.ApplyStyle(WorkbookStyleInputs.nextStyleInput(data)));
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(
                        targetSheet,
                        validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false)),
                    new MutationAction.SetDataValidation(nextDataValidationInput(data)));
            case 0x1 ->
                mutate(
                    nextRangeSelector(data, targetSheet, validRange),
                    new MutationAction.ClearDataValidations());
            case 0x2 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetConditionalFormatting(
                        nextConditionalFormattingInput(data, validRange)));
            default ->
                mutate(
                    nextRangeSelector(data, targetSheet, validRange),
                    new MutationAction.ClearConditionalFormatting());
          };
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(targetSheet, nextAutofilterRange(validRange)),
                    new MutationAction.SetAutofilter());
            case 0x1 ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.ClearAutofilter());
            case 0x2 ->
                mutate(
                    new MutationAction.SetTable(
                        nextTableInput(data, targetSheet, tableName, validRange)));
            case 0x3 ->
                mutate(
                    new TableSelector.ByNameOnSheet(tableName, targetSheet),
                    new MutationAction.DeleteTable());
            case 0x4 ->
                mutate(
                    new MutationAction.SetPivotTable(
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
                    new MutationAction.DeletePivotTable());
          };
      case 0x6 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new MutationAction.SetNamedRange(
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
                    new MutationAction.DeleteNamedRange());
          };
      case 0x7 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RowBandSelector.Insertion(
                        targetSheet, rowSpan.first(), rowSpan.last() - rowSpan.first() + 1),
                    new MutationAction.InsertRows());
            case 0x1 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.DeleteRows());
            case 0x2 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.ShiftRows(nextNonZeroDelta(data, 2)));
            case 0x3 ->
                mutate(
                    new ColumnBandSelector.Insertion(
                        targetSheet,
                        columnSpan.first(),
                        columnSpan.last() - columnSpan.first() + 1),
                    new MutationAction.InsertColumns());
            case 0x4 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.DeleteColumns());
            case 0x5 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.ShiftColumns(nextNonZeroDelta(data, 2)));
            case 0x6 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.SetRowVisibility(data.consumeBoolean()));
            case 0x7 ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.SetColumnVisibility(data.consumeBoolean()));
            case 0x8 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.GroupRows(data.consumeBoolean()));
            case 0x9 ->
                mutate(
                    new RowBandSelector.Span(targetSheet, rowSpan.first(), rowSpan.last()),
                    new MutationAction.UngroupRows());
            case 0xA ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.GroupColumns(data.consumeBoolean()));
            default ->
                mutate(
                    new ColumnBandSelector.Span(targetSheet, columnSpan.first(), columnSpan.last()),
                    new MutationAction.UngroupColumns());
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.AutoSizeColumns());
            case 0x1 ->
                mutate(
                    new WorkbookSelector.Current(),
                    new MutationAction.SetWorkbookProtection(
                        new WorkbookProtectionInput(true, false, false, null, null)));
            case 0x2 ->
                mutate(
                    new WorkbookSelector.Current(), new MutationAction.ClearWorkbookProtection());
            case 0x3 ->
                mutate(
                    new SheetSelector.ByName(targetSheet),
                    new MutationAction.SetSheetPresentation(SheetPresentationInput.defaults()));
            default ->
                mutate(new SheetSelector.ByName(targetSheet), new MutationAction.AutoSizeColumns());
          };
    };
  }
}
