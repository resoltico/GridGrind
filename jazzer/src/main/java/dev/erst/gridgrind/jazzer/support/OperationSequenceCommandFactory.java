package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceSelectorSupport.*;
import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelSheetPresentation;
import dev.erst.gridgrind.excel.ExcelWorkbookProtectionSettings;
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
            case 0x0 -> new WorkbookSheetCommand.CreateSheet(targetSheet);
            case 0x1 -> new WorkbookSheetCommand.RenameSheet(targetSheet, primarySheet + "Renamed");
            case 0x2 -> new WorkbookSheetCommand.DeleteSheet(targetSheet);
            case 0x3 -> new WorkbookSheetCommand.MoveSheet(targetSheet, data.consumeInt(0, 2));
            case 0x4 ->
                new WorkbookSheetCommand.CopySheet(
                    targetSheet, nextCopySheetName(targetSheet), nextExcelSheetCopyPosition(data));
            case 0x5 -> new WorkbookSheetCommand.SetActiveSheet(targetSheet);
            case 0x6 ->
                new WorkbookSheetCommand.SetSelectedSheets(
                    nextSelectedSheetNames(data, primarySheet, secondarySheet));
            case 0x7 ->
                new WorkbookSheetCommand.SetSheetVisibility(targetSheet, nextSheetVisibility(data));
            case 0x8 ->
                new WorkbookSheetCommand.SetSheetProtection(
                    targetSheet, nextSheetProtectionSettings(data));
            default -> new WorkbookSheetCommand.ClearSheetProtection(targetSheet);
          };
      case 0x1 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookStructureCommand.MergeCells(
                    targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
            case 0x1 ->
                new WorkbookStructureCommand.UnmergeCells(
                    targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
            case 0x2 ->
                new WorkbookStructureCommand.SetColumnWidth(
                    targetSheet,
                    columnSpan.first(),
                    columnSpan.last(),
                    data.consumeRegularDouble(1.0d, 20.0d));
            case 0x3 ->
                new WorkbookStructureCommand.SetRowHeight(
                    targetSheet,
                    rowSpan.first(),
                    rowSpan.last(),
                    data.consumeRegularDouble(5.0d, 40.0d));
            case 0x4 ->
                new WorkbookLayoutCommand.SetSheetPane(targetSheet, nextExcelSheetPane(data));
            case 0x5 ->
                new WorkbookLayoutCommand.SetSheetZoom(targetSheet, data.consumeInt(10, 400));
            case 0x6 ->
                new WorkbookLayoutCommand.SetPrintLayout(targetSheet, nextExcelPrintLayout(data));
            default -> new WorkbookLayoutCommand.ClearPrintLayout(targetSheet);
          };
      case 0x2 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookCellCommand.SetCell(
                    targetSheet,
                    FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                    FuzzDataDecoders.nextExcelCellValue(data));
            case 0x1 -> {
              String range = validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false);
              yield new WorkbookCellCommand.SetRange(
                  targetSheet, range, FuzzDataDecoders.nextExcelMatrix(data, 2, 2));
            }
            case 0x2 ->
                new WorkbookCellCommand.ClearRange(
                    targetSheet, FuzzDataDecoders.nextNonBlankRange(data, validRange));
            default ->
                new WorkbookCellCommand.AppendRow(
                    targetSheet,
                    List.of(
                        FuzzDataDecoders.nextExcelCellValue(data),
                        FuzzDataDecoders.nextExcelCellValue(data)));
          };
      case 0x3 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookAnnotationCommand.SetHyperlink(
                    targetSheet,
                    FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                    nextExcelHyperlink(data));
            case 0x1 ->
                new WorkbookAnnotationCommand.ClearHyperlink(
                    targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
            case 0x2 ->
                new WorkbookAnnotationCommand.SetComment(
                    targetSheet,
                    FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress),
                    nextExcelComment(data));
            case 0x3 ->
                new WorkbookAnnotationCommand.ClearComment(
                    targetSheet, FuzzDataDecoders.nextNonBlankCellAddress(data, validAddress));
            case 0x4 ->
                new WorkbookDrawingCommand.SetPicture(
                    targetSheet, nextExcelPictureDefinition(data));
            case 0x5 ->
                new WorkbookDrawingCommand.SetChart(targetSheet, nextExcelChartDefinition(data));
            case 0x6 ->
                new WorkbookDrawingCommand.SetShape(targetSheet, nextExcelShapeDefinition(data));
            case 0x7 ->
                new WorkbookDrawingCommand.SetEmbeddedObject(
                    targetSheet, nextExcelEmbeddedObjectDefinition(data));
            case 0x8 ->
                new WorkbookDrawingCommand.SetDrawingObjectAnchor(
                    targetSheet, nextDrawingObjectName(data), nextExcelDrawingAnchor(data));
            case 0x9 ->
                new WorkbookDrawingCommand.DeleteDrawingObject(
                    targetSheet, nextDrawingObjectName(data));
            default ->
                new WorkbookFormattingCommand.ApplyStyle(
                    targetSheet,
                    validRange ? "A1:B2" : FuzzDataDecoders.nextNonBlankRange(data, false),
                    FuzzDataDecoders.nextStyle(data));
          };
      case 0x4 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookFormattingCommand.SetDataValidation(
                    targetSheet,
                    validRange ? "A1:A4" : FuzzDataDecoders.nextNonBlankRange(data, false),
                    nextExcelDataValidationDefinition(data));
            case 0x1 ->
                new WorkbookFormattingCommand.ClearDataValidations(
                    targetSheet, nextExcelRangeSelection(data, validRange));
            case 0x2 ->
                new WorkbookFormattingCommand.SetConditionalFormatting(
                    targetSheet, nextExcelConditionalFormattingBlockDefinition(data, validRange));
            default ->
                new WorkbookFormattingCommand.ClearConditionalFormatting(
                    targetSheet, nextExcelRangeSelection(data, validRange));
          };
      case 0x5 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookTabularCommand.SetAutofilter(
                    targetSheet, nextAutofilterRange(validRange));
            case 0x1 -> new WorkbookTabularCommand.ClearAutofilter(targetSheet);
            case 0x2 ->
                new WorkbookTabularCommand.SetTable(
                    nextExcelTableDefinition(data, targetSheet, tableName, validRange));
            case 0x3 -> new WorkbookTabularCommand.DeleteTable(tableName, targetSheet);
            case 0x4 ->
                new WorkbookTabularCommand.SetPivotTable(
                    nextExcelPivotTableDefinition(
                        data,
                        targetSheet,
                        pivotTableName,
                        workbookNamedRange,
                        tableName,
                        validName,
                        validRange));
            default ->
                new WorkbookTabularCommand.DeletePivotTable(
                    validName ? pivotTableName : nextPivotTableName(data, false), targetSheet);
          };
      case 0x6 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookMetadataCommand.SetNamedRange(
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
                new WorkbookMetadataCommand.DeleteNamedRange(
                    validName ? namedRangeName : nextNamedRangeName(data, false),
                    data.consumeBoolean()
                        ? new ExcelNamedRangeScope.WorkbookScope()
                        : new ExcelNamedRangeScope.SheetScope(targetSheet));
          };
      case 0x7 ->
          switch (selectorSlot(selector)) {
            case 0x0 ->
                new WorkbookStructureCommand.InsertRows(
                    targetSheet, rowSpan.first(), rowSpan.last() - rowSpan.first() + 1);
            case 0x1 ->
                new WorkbookStructureCommand.DeleteRows(
                    targetSheet, new ExcelRowSpan(rowSpan.first(), rowSpan.last()));
            case 0x2 ->
                new WorkbookStructureCommand.ShiftRows(
                    targetSheet,
                    new ExcelRowSpan(rowSpan.first(), rowSpan.last()),
                    nextNonZeroDelta(data, 2));
            case 0x3 ->
                new WorkbookStructureCommand.InsertColumns(
                    targetSheet, columnSpan.first(), columnSpan.last() - columnSpan.first() + 1);
            case 0x4 ->
                new WorkbookStructureCommand.DeleteColumns(
                    targetSheet, new ExcelColumnSpan(columnSpan.first(), columnSpan.last()));
            case 0x5 ->
                new WorkbookStructureCommand.ShiftColumns(
                    targetSheet,
                    new ExcelColumnSpan(columnSpan.first(), columnSpan.last()),
                    nextNonZeroDelta(data, 2));
            case 0x6 ->
                new WorkbookStructureCommand.SetRowVisibility(
                    targetSheet,
                    new ExcelRowSpan(rowSpan.first(), rowSpan.last()),
                    data.consumeBoolean());
            case 0x7 ->
                new WorkbookStructureCommand.SetColumnVisibility(
                    targetSheet,
                    new ExcelColumnSpan(columnSpan.first(), columnSpan.last()),
                    data.consumeBoolean());
            case 0x8 ->
                new WorkbookStructureCommand.GroupRows(
                    targetSheet,
                    new ExcelRowSpan(rowSpan.first(), rowSpan.last()),
                    data.consumeBoolean());
            case 0x9 ->
                new WorkbookStructureCommand.UngroupRows(
                    targetSheet, new ExcelRowSpan(rowSpan.first(), rowSpan.last()));
            case 0xA ->
                new WorkbookStructureCommand.GroupColumns(
                    targetSheet,
                    new ExcelColumnSpan(columnSpan.first(), columnSpan.last()),
                    data.consumeBoolean());
            default ->
                new WorkbookStructureCommand.UngroupColumns(
                    targetSheet, new ExcelColumnSpan(columnSpan.first(), columnSpan.last()));
          };
      default ->
          switch (selectorSlot(selector)) {
            case 0x0 -> new WorkbookLayoutCommand.AutoSizeColumns(targetSheet);
            case 0x1 ->
                new WorkbookSheetCommand.SetWorkbookProtection(
                    new ExcelWorkbookProtectionSettings(true, false, false, null, null));
            case 0x2 -> new WorkbookSheetCommand.ClearWorkbookProtection();
            case 0x3 ->
                new WorkbookLayoutCommand.SetSheetPresentation(
                    targetSheet, ExcelSheetPresentation.defaults());
            default -> new WorkbookLayoutCommand.AutoSizeColumns(targetSheet);
          };
    };
  }
}
