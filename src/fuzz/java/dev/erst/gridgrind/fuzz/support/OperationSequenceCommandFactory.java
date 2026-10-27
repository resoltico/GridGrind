package dev.erst.gridgrind.fuzz.support;

import static dev.erst.gridgrind.fuzz.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.fuzz.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.ResolvedTarget;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.backend.xssf.*;
import java.util.List;
import java.util.Optional;

/** Builds bounded workbook-core command sequences for fuzz generation. */
final class OperationSequenceCommandFactory {
  private OperationSequenceCommandFactory() {}

  static ResolvedStep nextCommand(
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
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.EnsureSheet());
            case 0x1 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.RenameSheet(
                        new SheetName(primarySheet + "Renamed")));
            case 0x2 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.DeleteSheet());
            case 0x3 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.MoveSheet(data.consumeInt(0, 2)));
            case 0x4 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.CopySheet(
                        new SheetName(nextCopySheetName(targetSheet)),
                        nextExcelSheetCopyPosition(data)));
            case 0x5 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetActiveSheet());
            case 0x6 ->
                new ResolvedStep(
                    new SheetSelector.ByNames(
                        nextSelectedSheetNames(data, primarySheet, secondarySheet).stream()
                            .map(SheetName::new)
                            .toList()),
                    new WorkbookMutationAction.SetSelectedSheets());
            case 0x7 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetSheetVisibility(nextSheetVisibility(data)));
            case 0x8 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new ResolvedMutationAction.SetSheetProtection(
                        nextSheetProtectionSettings(data), Optional.empty()));
            default ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.ClearSheetProtection());
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new ResolvedStep(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(FuzzDataDecoders.nextNonBlankRange(data, validRange))),
                    new WorkbookMutationAction.MergeCells());
            case 0x1 ->
                new ResolvedStep(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(FuzzDataDecoders.nextNonBlankRange(data, validRange))),
                    new WorkbookMutationAction.UnmergeCells());
            case 0x2 ->
                new ResolvedStep(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.SetColumnWidth(
                        data.consumeRegularDouble(1.0d, 20.0d)));
            case 0x3 ->
                new ResolvedStep(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.SetRowHeight(
                        data.consumeRegularDouble(5.0d, 40.0d)));
            case 0x4 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetSheetPane(nextExcelSheetPane(data)));
            case 0x5 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetSheetZoom(data.consumeInt(10, 400)));
            case 0x6 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new ResolvedMutationAction.SetPrintLayout(nextExcelPrintLayout(data)));
            default ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.ClearPrintLayout());
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new ResolvedStep(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new ResolvedMutationAction.SetCell(FuzzDataDecoders.nextExcelCellValue(data)));
            case 0x1 -> {
              String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
              yield new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName(targetSheet), new RangeRef(range)),
                  new ResolvedMutationAction.SetRange(
                      FuzzDataDecoders.nextExcelMatrix(data, 2, 2)));
            }
            case 0x2 ->
                new ResolvedStep(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(FuzzDataDecoders.nextNonBlankRange(data, validRange))),
                    new CellMutationAction.ClearRange());
            default ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new ResolvedMutationAction.AppendRow(
                        List.of(
                            FuzzDataDecoders.nextExcelCellValue(data),
                            FuzzDataDecoders.nextExcelCellValue(data))));
          };
      case 0x3 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new ResolvedStep(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new CellMutationAction.SetHyperlink(nextExcelHyperlink(data)));
            case 0x1 ->
                new ResolvedStep(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new CellMutationAction.ClearHyperlink());
            case 0x2 ->
                new ResolvedStep(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new ResolvedMutationAction.SetComment(nextExcelComment(data)));
            case 0x3 ->
                new ResolvedStep(
                    new CellSelector.ByAddress(
                        new SheetName(targetSheet),
                        new CellAddress(
                            FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress))),
                    new CellMutationAction.ClearComment());
            case 0x4 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new ResolvedMutationAction.SetPicture(nextExcelPictureDefinition(data)));
            case 0x5 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new ResolvedMutationAction.SetChart(nextResolvedChart(data)));
            case 0x6 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new ResolvedMutationAction.SetShape(nextResolvedShape(data)));
            case 0x7 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new ResolvedMutationAction.SetEmbeddedObject(
                        nextExcelEmbeddedObjectDefinition(data)));
            case 0x8 ->
                new ResolvedStep(
                    new DrawingObjectSelector.ByName(
                        new SheetName(targetSheet), nextDrawingObjectName(data)),
                    new DrawingMutationAction.SetDrawingObjectAnchor(nextDrawingAnchorInput(data)));
            case 0x9 ->
                new ResolvedStep(
                    new DrawingObjectSelector.ByName(
                        new SheetName(targetSheet), nextDrawingObjectName(data)),
                    new DrawingMutationAction.DeleteDrawingObject());
            default ->
                new ResolvedStep(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(
                            validRange
                                ? "A1:B2"
                                : FuzzDataDecoders.nextNonBlankRange(data, false))),
                    new CellMutationAction.ApplyStyle(FuzzDataDecoders.nextStyle(data)));
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new ResolvedStep(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet),
                        new RangeRef(
                            validRange
                                ? "A1:A4"
                                : FuzzDataDecoders.nextNonBlankRange(data, false))),
                    new ResolvedMutationAction.SetDataValidation(
                        nextExcelDataValidationDefinition(data)));
            case 0x1 ->
                new ResolvedStep(
                    (dev.erst.gridgrind.core.plan.ResolvedTarget)
                        nextClearRangeTarget(data, validRange, targetSheet),
                    new StructuredMutationAction.ClearDataValidations());
            case 0x2 -> nextConditionalFormattingCommand(data, validRange, targetSheet);
            default ->
                new ResolvedStep(
                    (dev.erst.gridgrind.core.plan.ResolvedTarget)
                        nextClearRangeTarget(data, validRange, targetSheet),
                    new StructuredMutationAction.ClearConditionalFormatting());
          };
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new ResolvedStep(
                    new RangeSelector.ByRange(
                        new SheetName(targetSheet), new RangeRef(nextAutofilterRange(validRange))),
                    new StructuredMutationAction.SetAutofilter(List.of(), Optional.empty()));
            case 0x1 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new StructuredMutationAction.ClearAutofilter());
            case 0x2 ->
                setTableStep(nextExcelTableDefinition(data, targetSheet, tableName, validRange));
            case 0x3 ->
                new ResolvedStep(
                    new TableSelector.ByNameOnSheet(
                        new TableName(tableName), new SheetName(targetSheet)),
                    new StructuredMutationAction.DeleteTable());
            case 0x4 ->
                setPivotTableStep(
                    nextPivotTableSpec(
                        data,
                        targetSheet,
                        pivotTableName,
                        workbookNamedRange,
                        tableName,
                        validName,
                        validRange));
            default ->
                new ResolvedStep(
                    new PivotTableSelector.ByNameOnSheet(
                        validName ? pivotTableName : nextPivotTableName(data, false),
                        new SheetName(targetSheet)),
                    new StructuredMutationAction.DeletePivotTable());
          };
      case 0x6 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                setNamedRangeStep(
                    new StructuredMutationAction.SetNamedRange(
                        new dev.erst.gridgrind.core.vocabulary.DefinedName(
                            validName ? namedRangeName : nextNamedRangeName(data, false)),
                        data.consumeBoolean()
                            ? new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook()
                            : new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet(
                                new dev.erst.gridgrind.core.vocabulary.SheetName(targetSheet)),
                        NamedRangeTarget.range(
                            new dev.erst.gridgrind.core.vocabulary.SheetName(targetSheet),
                            new dev.erst.gridgrind.core.vocabulary.RangeRef(
                                validRange
                                    ? "A1:B2"
                                    : FuzzDataDecoders.nextNonBlankRange(data, false)))));
            default ->
                deleteNamedRangeStep(
                    validName ? namedRangeName : nextNamedRangeName(data, false),
                    data.consumeBoolean()
                        ? new NamedRangeScopeReport.Workbook()
                        : new NamedRangeScopeReport.Sheet(
                            new dev.erst.gridgrind.core.vocabulary.SheetLabel(targetSheet)));
          };
      case 0x7 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new ResolvedStep(
                    new RowBandSelector.Insertion(
                        new SheetName(targetSheet),
                        rowSpan.first(),
                        rowSpan.last() - rowSpan.first() + 1),
                    new WorkbookMutationAction.InsertRows());
            case 0x1 ->
                new ResolvedStep(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.DeleteRows());
            case 0x2 ->
                new ResolvedStep(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.ShiftRows(nextNonZeroDelta(data, 2)));
            case 0x3 ->
                new ResolvedStep(
                    new ColumnBandSelector.Insertion(
                        new SheetName(targetSheet),
                        columnSpan.first(),
                        columnSpan.last() - columnSpan.first() + 1),
                    new WorkbookMutationAction.InsertColumns());
            case 0x4 ->
                new ResolvedStep(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.DeleteColumns());
            case 0x5 ->
                new ResolvedStep(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.ShiftColumns(nextNonZeroDelta(data, 2)));
            case 0x6 ->
                new ResolvedStep(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.SetRowVisibility(data.consumeBoolean()));
            case 0x7 ->
                new ResolvedStep(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.SetColumnVisibility(data.consumeBoolean()));
            case 0x8 ->
                new ResolvedStep(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.GroupRows(data.consumeBoolean()));
            case 0x9 ->
                new ResolvedStep(
                    new RowBandSelector.Span(
                        new SheetName(targetSheet), rowSpan.first(), rowSpan.last()),
                    new WorkbookMutationAction.UngroupRows());
            case 0xA ->
                new ResolvedStep(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.GroupColumns(data.consumeBoolean()));
            default ->
                new ResolvedStep(
                    new ColumnBandSelector.Span(
                        new SheetName(targetSheet), columnSpan.first(), columnSpan.last()),
                    new WorkbookMutationAction.UngroupColumns());
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.AutoSizeColumns());
            case 0x1 ->
                new ResolvedStep(
                    new WorkbookSelector.Current(),
                    new ResolvedMutationAction.SetWorkbookProtection(
                        new ResolvedWorkbookProtection(
                            true, false, false, Optional.empty(), Optional.empty())));
            case 0x2 ->
                new ResolvedStep(
                    new WorkbookSelector.Current(),
                    new WorkbookMutationAction.ClearWorkbookProtection());
            case 0x3 ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.SetSheetPresentation(
                        SheetPresentationSpec.defaults()));
            default ->
                new ResolvedStep(
                    new SheetSelector.ByName(new SheetName(targetSheet)),
                    new WorkbookMutationAction.AutoSizeColumns());
          };
    };
  }

  static ResolvedStep setTableStep(dev.erst.gridgrind.core.model.table.ResolvedTable table) {
    return new ResolvedStep(
        new TableSelector.ByNameOnSheet(table.name(), table.sheetName()),
        new ResolvedMutationAction.SetTable(table));
  }

  static ResolvedStep setPivotTableStep(dev.erst.gridgrind.core.model.pivot.PivotTableSpec pivot) {
    return new ResolvedStep(
        new PivotTableSelector.ByNameOnSheet(pivot.name(), pivot.sheetName()),
        new StructuredMutationAction.SetPivotTable(pivot));
  }

  static ResolvedStep setNamedRangeStep(StructuredMutationAction.SetNamedRange action) {
    ResolvedTarget target =
        switch (action.scope()) {
          case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook _ ->
              new NamedRangeSelector.WorkbookScope(action.name());
          case dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Sheet scope ->
              new NamedRangeSelector.SheetScope(action.name(), scope.sheetName());
        };
    return new ResolvedStep(target, action);
  }

  static ResolvedStep deleteNamedRangeStep(
      String name, dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport scope) {
    ResolvedTarget target =
        switch (scope) {
          case dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport.Workbook _ ->
              new NamedRangeSelector.WorkbookScope(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName(name));
          case dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport.Sheet sheet ->
              new NamedRangeSelector.SheetScope(
                  new dev.erst.gridgrind.core.vocabulary.DefinedName(name),
                  new SheetName(sheet.sheetName().value()));
        };
    return new ResolvedStep(target, new StructuredMutationAction.DeleteNamedRange());
  }
}
