package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.mutate;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.Optional;

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
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.EnsureSheet());
            case 0x1 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.RenameSheet(
                        new SheetName(primarySheet + "Renamed")));
            case 0x2 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.DeleteSheet());
            case 0x3 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.MoveSheet(data.consumeInt(0, 2)));
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.CopySheet(
                        new SheetName(nextCopySheetName(targetSheet)),
                        nextSheetCopyPosition(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetActiveSheet());
            case 0x6 ->
                mutate(
                    new SheetSelector.ByNames(
                        nextSelectedSheetNames(data, primarySheet, secondarySheet).stream()
                            .map(SheetName::new)
                            .toList()),
                    new WorkbookMutationAction.SetSelectedSheets());
            case 0x7 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetSheetVisibility(
                        nextProtocolSheetVisibility(data)));
            case 0x8 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetSheetProtection(
                        nextProtocolSheetProtectionSettings(data)));
            default ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.ClearSheetProtection());
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(FuzzDataDecoders.nextNonBlankRange(data, validRange))),
                    new WorkbookMutationAction.MergeCells());
            case 0x1 ->
                mutate(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(FuzzDataDecoders.nextNonBlankRange(data, validRange))),
                    new WorkbookMutationAction.UnmergeCells());
            case 0x2 ->
                mutate(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.SetColumnWidth(
                        data.consumeRegularDouble(1.0d, 20.0d)));
            case 0x3 ->
                mutate(
                    nextRowHeightTarget(data, targetSheet, rowSpan),
                    new WorkbookMutationAction.SetRowHeight(
                        data.consumeRegularDouble(5.0d, 40.0d)));
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetSheetPane(nextPaneInput(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetSheetZoom(data.consumeInt(10, 400)));
            case 0x6 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetPrintLayout(nextPrintLayoutInput(data)));
            default ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.ClearPrintLayout());
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new CellMutationAction.SetCell(FuzzDataDecoders.nextCellInput(data)));
            case 0x1 -> {
              String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
              yield mutate(
                  new RangeSelector.ByRange(new SheetName(targetSheet), new RangeRef(range)),
                  new CellMutationAction.SetRange(FuzzDataDecoders.nextProtocolGrid(data, 2, 2)));
            }
            case 0x2 ->
                mutate(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(FuzzDataDecoders.nextNonBlankRange(data, validRange))),
                    new CellMutationAction.ClearRange());
            default ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new CellMutationAction.AppendRow(FuzzDataDecoders.nextProtocolRow(data, 2)));
          };
      case 0x3 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new CellMutationAction.SetHyperlink(nextHyperlinkTarget(data)));
            case 0x1 ->
                mutate(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new CellMutationAction.ClearHyperlink());
            case 0x2 ->
                mutate(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new CellMutationAction.SetComment(nextCommentInput(data)));
            case 0x3 ->
                mutate(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new CellMutationAction.ClearComment());
            case 0x4 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new DrawingMutationAction.SetPicture(nextPictureInput(data)));
            case 0x5 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new DrawingMutationAction.SetChart(nextChartInput(data)));
            case 0x6 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new DrawingMutationAction.SetShape(nextShapeInput(data)));
            case 0x7 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new DrawingMutationAction.SetEmbeddedObject(nextEmbeddedObjectInput(data)));
            case 0x8 ->
                mutate(
                    new DrawingObjectSelector.ByName(
                        new SheetName(targetSheet), nextDrawingObjectName(data)),
                    new DrawingMutationAction.SetDrawingObjectAnchor(nextDrawingAnchorInput(data)));
            case 0x9 ->
                mutate(
                    new DrawingObjectSelector.ByName(
                        new SheetName(targetSheet), nextDrawingObjectName(data)),
                    new DrawingMutationAction.DeleteDrawingObject());
            default ->
                mutate(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(
                            validRange
                                ? "A1:B2"
                                : FuzzDataDecoders.nextNonBlankRange(data, false))),
                    new CellMutationAction.ApplyStyle(WorkbookStyleInputs.nextStyleInput(data)));
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(
                            validRange
                                ? "A1:A4"
                                : FuzzDataDecoders.nextNonBlankRange(data, false))),
                    new StructuredMutationAction.SetDataValidation(nextDataValidationInput(data)));
            case 0x1 ->
                mutate(
                    nextRangeSelector(data, targetSheet, validRange),
                    new StructuredMutationAction.ClearDataValidations());
            case 0x2 ->
                mutate(
                    nextConditionalFormattingTarget(data, targetSheet, validRange),
                    new StructuredMutationAction.SetConditionalFormatting(
                        nextConditionalFormattingInput(data)));
            default ->
                mutate(
                    nextRangeSelector(data, targetSheet, validRange),
                    new StructuredMutationAction.ClearConditionalFormatting());
          };
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet), new RangeRef(nextAutofilterRange(validRange))),
                    new StructuredMutationAction.SetAutofilter());
            case 0x1 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new StructuredMutationAction.ClearAutofilter());
            case 0x2 ->
                mutate(
                    new StructuredMutationAction.SetTable(
                        nextTableInput(data, targetSheet, tableName, validRange)));
            case 0x3 ->
                mutate(
                    new TableSelector.ByNameOnSheet(
                        new TableName(tableName), new SheetName(targetSheet)),
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
                        validName ? pivotTableName : nextPivotTableName(data, false),
                        new SheetName(targetSheet)),
                    new StructuredMutationAction.DeletePivotTable());
          };
      case 0x6 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new StructuredMutationAction.SetNamedRange(
                        new dev.erst.gridgrind.core.vocabulary.DefinedName(
                            validName ? namedRangeName : nextNamedRangeName(data, false)),
                        data.consumeBoolean()
                            ? new NamedRangeScope.Workbook()
                            : new NamedRangeScope.Sheet(new SheetName(targetSheet)),
                        NamedRangeTarget.range(
                            new SheetName(targetSheet),
                            new RangeRef(
                                validRange
                                    ? "A1:B2"
                                    : FuzzDataDecoders.nextNonBlankRange(data, false)))));
            default ->
                mutate(
                    data.consumeBoolean()
                        ? new NamedRangeSelector.WorkbookScope(
                            new dev.erst.gridgrind.core.vocabulary.DefinedName(
                                validName ? namedRangeName : nextNamedRangeName(data, false)))
                        : new NamedRangeSelector.SheetScope(
                            new dev.erst.gridgrind.core.vocabulary.DefinedName(
                                validName ? namedRangeName : nextNamedRangeName(data, false)),
                            new SheetName(targetSheet)),
                    new StructuredMutationAction.DeleteNamedRange());
          };
      case 0x7 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new RowBandSelector.Insertion(
                        new SheetName(targetSheet),
                        rowSpan.first(),
                        rowSpan.last() - rowSpan.first() + 1),
                    new WorkbookMutationAction.InsertRows());
            case 0x1 ->
                mutate(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.DeleteRows());
            case 0x2 ->
                mutate(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.ShiftRows(nextNonZeroDelta(data, 2)));
            case 0x3 ->
                mutate(
                    new ColumnBandSelector.Insertion(
                        new SheetName(targetSheet),
                        columnSpan.first(),
                        columnSpan.last() - columnSpan.first() + 1),
                    new WorkbookMutationAction.InsertColumns());
            case 0x4 ->
                mutate(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.DeleteColumns());
            case 0x5 ->
                mutate(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.ShiftColumns(nextNonZeroDelta(data, 2)));
            case 0x6 ->
                mutate(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.SetRowVisibility(data.consumeBoolean()));
            case 0x7 ->
                mutate(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.SetColumnVisibility(data.consumeBoolean()));
            case 0x8 ->
                mutate(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.GroupRows(data.consumeBoolean()));
            case 0x9 ->
                mutate(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.UngroupRows());
            case 0xA ->
                mutate(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.GroupColumns(data.consumeBoolean()));
            default ->
                mutate(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.UngroupColumns());
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.AutoSizeColumns());
            case 0x1 ->
                mutate(
                    new WorkbookSelector.Current(),
                    new WorkbookMutationAction.SetWorkbookProtection(
                        new WorkbookProtectionSpec(
                            true, false, false, Optional.empty(), Optional.empty())));
            case 0x2 ->
                mutate(
                    new WorkbookSelector.Current(),
                    new WorkbookMutationAction.ClearWorkbookProtection());
            case 0x3 ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetSheetPresentation(
                        SheetPresentationSpec.defaults()));
            default ->
                mutate(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.AutoSizeColumns());
          };
    };
  }
}
