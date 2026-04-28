package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.List;

/** Builds bounded workbook-core command sequences for Jazzer generation. */
final class OperationSequenceCommandFactory {
  private OperationSequenceCommandFactory() {}

  static WorkbookCommand nextCommand(
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
            case 0x0 -> new WorkbookCommand.CreateSheet(targetSheet);
            case 0x1 -> new WorkbookCommand.RenameSheet(targetSheet, primarySheet + "Renamed");
            case 0x2 -> new WorkbookCommand.DeleteSheet(targetSheet);
            case 0x3 -> new WorkbookCommand.MoveSheet(targetSheet, data.consumeInt(0, 2));
            case 0x4 ->
                new WorkbookCommand.CopySheet(
                    targetSheet, nextCopySheetName(targetSheet), nextExcelSheetCopyPosition(data));
            case 0x5 -> new WorkbookCommand.SetActiveSheet(targetSheet);
            case 0x6 ->
                new WorkbookCommand.SetSelectedSheets(
                    nextSelectedSheetNames(data, primarySheet, secondarySheet));
            case 0x7 ->
                new WorkbookCommand.SetSheetVisibility(targetSheet, nextSheetVisibility(data));
            case 0x8 ->
                new WorkbookCommand.SetSheetProtection(
                    targetSheet, nextSheetProtectionSettings(data));
            default -> new WorkbookCommand.ClearSheetProtection(targetSheet);
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.MergeCells(
                    targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
            case 0x1 ->
                new WorkbookCommand.UnmergeCells(
                    targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
            case 0x2 ->
                new WorkbookCommand.SetColumnWidth(
                    targetSheet,
                    columnSpan.first(),
                    columnSpan.last(),
                    data.consumeRegularDouble(1.0d, 20.0d));
            case 0x3 ->
                new WorkbookCommand.SetRowHeight(
                    targetSheet,
                    rowSpan.first(),
                    rowSpan.last(),
                    data.consumeRegularDouble(5.0d, 40.0d));
            case 0x4 -> new WorkbookCommand.SetSheetPane(targetSheet, nextExcelSheetPane(data));
            case 0x5 -> new WorkbookCommand.SetSheetZoom(targetSheet, data.consumeInt(10, 400));
            case 0x6 -> new WorkbookCommand.SetPrintLayout(targetSheet, nextExcelPrintLayout(data));
            default -> new WorkbookCommand.ClearPrintLayout(targetSheet);
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.SetCell(
                    targetSheet,
                    FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                    FuzzDataDecoders.nextExcelCellValue(data));
            case 0x1 -> {
              String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
              yield new WorkbookCommand.SetRange(
                  targetSheet, range, FuzzDataDecoders.nextExcelMatrix(data, 2, 2));
            }
            case 0x2 ->
                new WorkbookCommand.ClearRange(
                    targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
            default ->
                new WorkbookCommand.AppendRow(
                    targetSheet,
                    List.of(
                        FuzzDataDecoders.nextExcelCellValue(data),
                        FuzzDataDecoders.nextExcelCellValue(data)));
          };
      case 0x3 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.SetHyperlink(
                    targetSheet,
                    FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                    nextExcelHyperlink(data));
            case 0x1 ->
                new WorkbookCommand.ClearHyperlink(
                    targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
            case 0x2 ->
                new WorkbookCommand.SetComment(
                    targetSheet,
                    FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                    nextExcelComment(data));
            case 0x3 ->
                new WorkbookCommand.ClearComment(
                    targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
            case 0x4 ->
                new WorkbookCommand.SetPicture(targetSheet, nextExcelPictureDefinition(data));
            case 0x5 -> new WorkbookCommand.SetChart(targetSheet, nextExcelChartDefinition(data));
            case 0x6 -> new WorkbookCommand.SetShape(targetSheet, nextExcelShapeDefinition(data));
            case 0x7 ->
                new WorkbookCommand.SetEmbeddedObject(
                    targetSheet, nextExcelEmbeddedObjectDefinition(data));
            case 0x8 ->
                new WorkbookCommand.SetDrawingObjectAnchor(
                    targetSheet, nextDrawingObjectName(data), nextExcelDrawingAnchor(data));
            case 0x9 ->
                new WorkbookCommand.DeleteDrawingObject(targetSheet, nextDrawingObjectName(data));
            default ->
                new WorkbookCommand.ApplyStyle(
                    targetSheet,
                    validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false),
                    FuzzDataDecoders.nextStyle(data));
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.SetDataValidation(
                    targetSheet,
                    validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false),
                    nextExcelDataValidationDefinition(data));
            case 0x1 ->
                new WorkbookCommand.ClearDataValidations(
                    targetSheet, nextExcelRangeSelection(data, validRange));
            case 0x2 ->
                new WorkbookCommand.SetConditionalFormatting(
                    targetSheet, nextExcelConditionalFormattingBlockDefinition(data, validRange));
            default ->
                new WorkbookCommand.ClearConditionalFormatting(
                    targetSheet, nextExcelRangeSelection(data, validRange));
          };
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.SetAutofilter(targetSheet, nextAutofilterRange(validRange));
            case 0x1 -> new WorkbookCommand.ClearAutofilter(targetSheet);
            case 0x2 ->
                new WorkbookCommand.SetTable(
                    nextExcelTableDefinition(data, targetSheet, tableName, validRange));
            case 0x3 -> new WorkbookCommand.DeleteTable(tableName, targetSheet);
            case 0x4 ->
                new WorkbookCommand.SetPivotTable(
                    nextExcelPivotTableDefinition(
                        data,
                        targetSheet,
                        pivotTableName,
                        workbookNamedRange,
                        tableName,
                        validName,
                        validRange));
            default ->
                new WorkbookCommand.DeletePivotTable(
                    validName ? pivotTableName : nextPivotTableName(data, false), targetSheet);
          };
      case 0x6 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.SetNamedRange(
                    new ExcelNamedRangeDefinition(
                        validName ? namedRangeName : nextNamedRangeName(data, false),
                        data.consumeBoolean()
                            ? new ExcelNamedRangeScope.WorkbookScope()
                            : new ExcelNamedRangeScope.SheetScope(targetSheet),
                        new ExcelNamedRangeTarget(
                            targetSheet,
                            validRange
                                ? "A1:B2"
                                : FuzzDataDecoders.nextNonBlankRange(data, false))));
            default ->
                new WorkbookCommand.DeleteNamedRange(
                    validName ? namedRangeName : nextNamedRangeName(data, false),
                    data.consumeBoolean()
                        ? new ExcelNamedRangeScope.WorkbookScope()
                        : new ExcelNamedRangeScope.SheetScope(targetSheet));
          };
      case 0x7 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCommand.InsertRows(
                    targetSheet, rowSpan.first(), rowSpan.last() - rowSpan.first() + 1);
            case 0x1 ->
                new WorkbookCommand.DeleteRows(
                    targetSheet, new ExcelRowSpan(rowSpan.first(), rowSpan.last()));
            case 0x2 ->
                new WorkbookCommand.ShiftRows(
                    targetSheet,
                    new ExcelRowSpan(rowSpan.first(), rowSpan.last()),
                    nextNonZeroDelta(data, 2));
            case 0x3 ->
                new WorkbookCommand.InsertColumns(
                    targetSheet, columnSpan.first(), columnSpan.last() - columnSpan.first() + 1);
            case 0x4 ->
                new WorkbookCommand.DeleteColumns(
                    targetSheet, new ExcelColumnSpan(columnSpan.first(), columnSpan.last()));
            case 0x5 ->
                new WorkbookCommand.ShiftColumns(
                    targetSheet,
                    new ExcelColumnSpan(columnSpan.first(), columnSpan.last()),
                    nextNonZeroDelta(data, 2));
            case 0x6 ->
                new WorkbookCommand.SetRowVisibility(
                    targetSheet,
                    new ExcelRowSpan(rowSpan.first(), rowSpan.last()),
                    data.consumeBoolean());
            case 0x7 ->
                new WorkbookCommand.SetColumnVisibility(
                    targetSheet,
                    new ExcelColumnSpan(columnSpan.first(), columnSpan.last()),
                    data.consumeBoolean());
            case 0x8 ->
                new WorkbookCommand.GroupRows(
                    targetSheet,
                    new ExcelRowSpan(rowSpan.first(), rowSpan.last()),
                    data.consumeBoolean());
            case 0x9 ->
                new WorkbookCommand.UngroupRows(
                    targetSheet, new ExcelRowSpan(rowSpan.first(), rowSpan.last()));
            case 0xA ->
                new WorkbookCommand.GroupColumns(
                    targetSheet,
                    new ExcelColumnSpan(columnSpan.first(), columnSpan.last()),
                    data.consumeBoolean());
            default ->
                new WorkbookCommand.UngroupColumns(
                    targetSheet, new ExcelColumnSpan(columnSpan.first(), columnSpan.last()));
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 -> new WorkbookCommand.AutoSizeColumns(targetSheet);
            case 0x1 ->
                new WorkbookCommand.SetWorkbookProtection(
                    new ExcelWorkbookProtectionSettings(true, false, false, null, null));
            case 0x2 -> new WorkbookCommand.ClearWorkbookProtection();
            case 0x3 ->
                new WorkbookCommand.SetSheetPresentation(
                    targetSheet, ExcelSheetPresentation.defaults());
            default -> new WorkbookCommand.AutoSizeColumns(targetSheet);
          };
    };
  }
}
