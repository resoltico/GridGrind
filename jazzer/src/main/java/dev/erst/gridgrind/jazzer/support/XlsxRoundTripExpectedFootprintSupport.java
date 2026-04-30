package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookAnnotationCommand;
import dev.erst.gridgrind.excel.WorkbookCellCommand;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookDrawingCommand;
import dev.erst.gridgrind.excel.WorkbookFormattingCommand;
import dev.erst.gridgrind.excel.WorkbookLayoutCommand;
import dev.erst.gridgrind.excel.WorkbookMetadataCommand;
import dev.erst.gridgrind.excel.WorkbookSheetCommand;
import dev.erst.gridgrind.excel.WorkbookStructureCommand;
import dev.erst.gridgrind.excel.WorkbookTabularCommand;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;
import org.apache.poi.ss.util.CellRangeAddress;

/** Owns candidate-cell footprint tracking for `.xlsx` round-trip expectations. */
final class XlsxRoundTripExpectedFootprintSupport {
  private XlsxRoundTripExpectedFootprintSupport() {}

  static XlsxRoundTripExpectedStateSupport.ExpectedWorkbookFootprint expectedWorkbookFootprint(
      List<WorkbookCommand> commands) {
    LinkedHashMap<String, Set<XlsxRoundTripExpectedStateSupport.CellCoordinate>>
        candidateCoordinatesBySheet = new LinkedHashMap<>();
    LinkedHashMap<String, Set<XlsxRoundTripExpectedStateSupport.CellCoordinate>>
        valueBearingCoordinatesBySheet = new LinkedHashMap<>();

    for (WorkbookCommand command : commands) {
      switch (command) {
        case WorkbookSheetCommand.CreateSheet createSheet ->
            candidateCoordinatesBySheet.computeIfAbsent(
                createSheet.sheetName(), sheetKey -> new LinkedHashSet<>());
        case WorkbookSheetCommand.RenameSheet renameSheet -> {
          renameExpectedSheetState(
              candidateCoordinatesBySheet, renameSheet.sheetName(), renameSheet.newSheetName());
          renameExpectedSheetState(
              valueBearingCoordinatesBySheet, renameSheet.sheetName(), renameSheet.newSheetName());
        }
        case WorkbookSheetCommand.DeleteSheet deleteSheet -> {
          candidateCoordinatesBySheet.remove(deleteSheet.sheetName());
          valueBearingCoordinatesBySheet.remove(deleteSheet.sheetName());
        }
        case WorkbookSheetCommand.MoveSheet _ -> {
          // Sheet order does not affect candidate or append-location tracking.
        }
        case WorkbookSheetCommand.CopySheet copySheet -> {
          copyExpectedSheetState(
              candidateCoordinatesBySheet, copySheet.sourceSheetName(), copySheet.newSheetName());
          copyExpectedSheetState(
              valueBearingCoordinatesBySheet,
              copySheet.sourceSheetName(),
              copySheet.newSheetName());
        }
        case WorkbookSheetCommand.SetActiveSheet _ -> {
          // Active sheet state does not affect cell-level round-trip expectations.
        }
        case WorkbookSheetCommand.SetSelectedSheets _ -> {
          // Selected sheet state does not affect cell-level round-trip expectations.
        }
        case WorkbookSheetCommand.SetSheetVisibility _ -> {
          // Visibility state does not affect cell-level round-trip expectations.
        }
        case WorkbookSheetCommand.SetSheetProtection _ -> {
          // Sheet protection is tracked independently from cell-level expectations.
        }
        case WorkbookSheetCommand.ClearSheetProtection _ -> {
          // Sheet protection is tracked independently from cell-level expectations.
        }
        case WorkbookSheetCommand.SetWorkbookProtection _ -> {
          // Workbook protection is tracked independently from cell-level expectations.
        }
        case WorkbookSheetCommand.ClearWorkbookProtection _ -> {
          // Workbook protection is tracked independently from cell-level expectations.
        }
        case WorkbookStructureCommand.MergeCells _ -> {
          // Merge state does not add new candidate cells.
        }
        case WorkbookStructureCommand.UnmergeCells _ -> {
          // Unmerge state does not add new candidate cells.
        }
        case WorkbookStructureCommand.SetColumnWidth _ -> {
          // Column width does not affect cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.SetRowHeight _ -> {
          // Row height does not affect cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.InsertRows _ -> {
          // Structural row state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.DeleteRows _ -> {
          // Structural row state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.ShiftRows _ -> {
          // Structural row state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.InsertColumns _ -> {
          // Structural column state is tracked independently from cell-level round-trip
          // expectations.
        }
        case WorkbookStructureCommand.DeleteColumns _ -> {
          // Structural column state is tracked independently from cell-level round-trip
          // expectations.
        }
        case WorkbookStructureCommand.ShiftColumns _ -> {
          // Structural column state is tracked independently from cell-level round-trip
          // expectations.
        }
        case WorkbookStructureCommand.SetRowVisibility _ -> {
          // Row visibility is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.SetColumnVisibility _ -> {
          // Column visibility is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.GroupRows _ -> {
          // Row outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.UngroupRows _ -> {
          // Row outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.GroupColumns _ -> {
          // Column outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookStructureCommand.UngroupColumns _ -> {
          // Column outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookLayoutCommand.SetSheetPane _ -> {
          // Pane state does not affect cell-level round-trip expectations.
        }
        case WorkbookLayoutCommand.SetSheetZoom _ -> {
          // Zoom state does not affect cell-level round-trip expectations.
        }
        case WorkbookLayoutCommand.SetSheetPresentation _ -> {
          // Sheet presentation does not affect cell-level round-trip expectations.
        }
        case WorkbookLayoutCommand.SetPrintLayout _ -> {
          // Print layout does not affect cell-level round-trip expectations.
        }
        case WorkbookLayoutCommand.ClearPrintLayout _ -> {
          // Print layout clearing does not affect cell-level round-trip expectations.
        }
        case WorkbookCellCommand.SetCell setCell -> {
          XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate =
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(setCell.address());
          recordCandidate(candidateCoordinatesBySheet, setCell.sheetName(), coordinate);
          recordValueBearing(
              valueBearingCoordinatesBySheet,
              setCell.sheetName(),
              coordinate,
              isValueBearing(setCell.value()));
        }
        case WorkbookCellCommand.SetRange setRange -> {
          CellRangeAddress range = CellRangeAddress.valueOf(setRange.range());
          for (int rowOffset = 0; rowOffset < setRange.rows().size(); rowOffset++) {
            for (int columnOffset = 0;
                columnOffset < setRange.rows().get(rowOffset).size();
                columnOffset++) {
              XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate =
                  new XlsxRoundTripExpectedStateSupport.CellCoordinate(
                      range.getFirstRow() + rowOffset, range.getFirstColumn() + columnOffset);
              recordCandidate(candidateCoordinatesBySheet, setRange.sheetName(), coordinate);
              recordValueBearing(
                  valueBearingCoordinatesBySheet,
                  setRange.sheetName(),
                  coordinate,
                  isValueBearing(setRange.rows().get(rowOffset).get(columnOffset)));
            }
          }
        }
        case WorkbookCellCommand.ClearRange clearRange ->
            forEachCell(
                clearRange.range(),
                coordinate -> {
                  recordCandidate(candidateCoordinatesBySheet, clearRange.sheetName(), coordinate);
                  recordValueBearing(
                      valueBearingCoordinatesBySheet, clearRange.sheetName(), coordinate, false);
                });
        case WorkbookCellCommand.SetArrayFormula setArrayFormula ->
            forEachCell(
                setArrayFormula.range(),
                coordinate -> {
                  recordCandidate(
                      candidateCoordinatesBySheet, setArrayFormula.sheetName(), coordinate);
                  recordValueBearing(
                      valueBearingCoordinatesBySheet,
                      setArrayFormula.sheetName(),
                      coordinate,
                      true);
                });
        case WorkbookCellCommand.ClearArrayFormula clearArrayFormula -> {
          XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate =
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                  clearArrayFormula.address());
          recordCandidate(candidateCoordinatesBySheet, clearArrayFormula.sheetName(), coordinate);
          recordValueBearing(
              valueBearingCoordinatesBySheet, clearArrayFormula.sheetName(), coordinate, false);
        }
        case WorkbookMetadataCommand.ImportCustomXmlMapping _ -> {
          // Custom-XML imports mutate workbook content outside the fuzz model's candidate-cell set.
        }
        case WorkbookAnnotationCommand.SetHyperlink setHyperlink ->
            recordCandidate(
                candidateCoordinatesBySheet,
                setHyperlink.sheetName(),
                XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                    setHyperlink.address()));
        case WorkbookAnnotationCommand.ClearHyperlink clearHyperlink ->
            recordCandidate(
                candidateCoordinatesBySheet,
                clearHyperlink.sheetName(),
                XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                    clearHyperlink.address()));
        case WorkbookAnnotationCommand.SetComment setComment ->
            recordCandidate(
                candidateCoordinatesBySheet,
                setComment.sheetName(),
                XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(setComment.address()));
        case WorkbookAnnotationCommand.ClearComment clearComment ->
            recordCandidate(
                candidateCoordinatesBySheet,
                clearComment.sheetName(),
                XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                    clearComment.address()));
        case WorkbookDrawingCommand.SetPicture _ -> {
          // Drawing objects are validated through workbook readability, not candidate cell grids.
        }
        case WorkbookDrawingCommand.SetSignatureLine _ -> {
          // Signature lines are validated through workbook readability, not candidate cell grids.
        }
        case WorkbookDrawingCommand.SetShape _ -> {
          // Drawing objects are validated through workbook readability, not candidate cell grids.
        }
        case WorkbookDrawingCommand.SetEmbeddedObject _ -> {
          // Drawing objects are validated through workbook readability, not candidate cell grids.
        }
        case WorkbookDrawingCommand.SetChart _ -> {
          // Charts are validated through workbook-level drawing and chart snapshots.
        }
        case WorkbookTabularCommand.SetPivotTable _ -> {
          // Pivot tables are validated through workbook-level pivot snapshots.
        }
        case WorkbookDrawingCommand.SetDrawingObjectAnchor _ -> {
          // Anchor mutation changes drawing geometry only.
        }
        case WorkbookDrawingCommand.DeleteDrawingObject _ -> {
          // Drawing deletion does not alter candidate cell tracking.
        }
        case WorkbookFormattingCommand.ApplyStyle applyStyle ->
            forEachCell(
                applyStyle.range(),
                coordinate ->
                    recordCandidate(
                        candidateCoordinatesBySheet, applyStyle.sheetName(), coordinate));
        case WorkbookFormattingCommand.SetDataValidation _ -> {
          // Data validations are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookFormattingCommand.ClearDataValidations _ -> {
          // Data validations are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookFormattingCommand.SetConditionalFormatting _ -> {
          // Conditional formatting is tracked independently from cell-style and metadata
          // expectations.
        }
        case WorkbookFormattingCommand.ClearConditionalFormatting _ -> {
          // Conditional formatting is tracked independently from cell-style and metadata
          // expectations.
        }
        case WorkbookTabularCommand.SetAutofilter _ -> {
          // Autofilters are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookTabularCommand.ClearAutofilter _ -> {
          // Autofilters are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookTabularCommand.SetTable _ -> {
          // Tables are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookTabularCommand.DeleteTable _ -> {
          // Tables are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookTabularCommand.DeletePivotTable _ -> {
          // Pivot tables are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookMetadataCommand.SetNamedRange _ -> {
          // Named ranges are derived directly from the applied workbook state.
        }
        case WorkbookMetadataCommand.DeleteNamedRange _ -> {
          // Named ranges are derived directly from the applied workbook state.
        }
        case WorkbookCellCommand.AppendRow appendRow -> {
          int rowIndex = nextAppendRowIndex(valueBearingCoordinatesBySheet, appendRow.sheetName());
          for (int columnIndex = 0; columnIndex < appendRow.values().size(); columnIndex++) {
            XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate =
                new XlsxRoundTripExpectedStateSupport.CellCoordinate(rowIndex, columnIndex);
            recordCandidate(candidateCoordinatesBySheet, appendRow.sheetName(), coordinate);
            recordValueBearing(
                valueBearingCoordinatesBySheet,
                appendRow.sheetName(),
                coordinate,
                isValueBearing(appendRow.values().get(columnIndex)));
          }
        }
        case WorkbookLayoutCommand.AutoSizeColumns _ -> {
          // Auto-sizing does not add new candidate cells.
        }
      }
    }

    LinkedHashMap<String, List<XlsxRoundTripExpectedStateSupport.CellCoordinate>>
        snapshotCoordinatesBySheet = new LinkedHashMap<>();
    candidateCoordinatesBySheet.forEach(
        (sheetName, coordinates) -> {
          if (!coordinates.isEmpty()) {
            snapshotCoordinatesBySheet.put(sheetName, sortedCoordinates(coordinates));
          }
        });
    return new XlsxRoundTripExpectedStateSupport.ExpectedWorkbookFootprint(
        Map.copyOf(snapshotCoordinatesBySheet));
  }

  static Map<String, List<ExcelCellSnapshot>> expectedCellSnapshots(
      ExcelWorkbook workbook,
      XlsxRoundTripExpectedStateSupport.ExpectedWorkbookFootprint footprint) {
    LinkedHashMap<String, List<ExcelCellSnapshot>> snapshotsBySheet = new LinkedHashMap<>();
    for (Map.Entry<String, List<XlsxRoundTripExpectedStateSupport.CellCoordinate>> entry :
        footprint.candidateCoordinatesBySheet().entrySet()) {
      List<String> addresses =
          entry.getValue().stream()
              .map(XlsxRoundTripExpectedStateSupport.CellCoordinate::a1Address)
              .toList();
      List<ExcelCellSnapshot> snapshots = workbook.sheet(entry.getKey()).snapshotCells(addresses);
      if (!snapshots.isEmpty()) {
        snapshotsBySheet.put(entry.getKey(), snapshots);
      }
    }
    return Map.copyOf(snapshotsBySheet);
  }

  static <T> void renameExpectedSheetState(
      Map<String, T> valuesBySheet, String sheetName, String newSheetName) {
    T values = valuesBySheet.remove(sheetName);
    if (values != null) {
      valuesBySheet.put(newSheetName, values);
    }
  }

  static <T> void copyExpectedSheetState(
      Map<String, T> valuesBySheet, String sourceSheetName, String newSheetName) {
    T values = valuesBySheet.get(sourceSheetName);
    if (values != null) {
      valuesBySheet.put(newSheetName, values);
    }
  }

  static void recordCandidate(
      Map<String, Set<XlsxRoundTripExpectedStateSupport.CellCoordinate>>
          candidateCoordinatesBySheet,
      String sheetName,
      XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate) {
    candidateCoordinatesBySheet
        .computeIfAbsent(sheetName, sheetKey -> new LinkedHashSet<>())
        .add(coordinate);
  }

  static void recordValueBearing(
      Map<String, Set<XlsxRoundTripExpectedStateSupport.CellCoordinate>>
          valueBearingCoordinatesBySheet,
      String sheetName,
      XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate,
      boolean valueBearing) {
    Set<XlsxRoundTripExpectedStateSupport.CellCoordinate> coordinates =
        valueBearingCoordinatesBySheet.computeIfAbsent(
            sheetName, sheetKey -> new LinkedHashSet<>());
    if (valueBearing) {
      coordinates.add(coordinate);
    } else {
      coordinates.remove(coordinate);
      if (coordinates.isEmpty()) {
        valueBearingCoordinatesBySheet.remove(sheetName);
      }
    }
  }

  static int nextAppendRowIndex(
      Map<String, Set<XlsxRoundTripExpectedStateSupport.CellCoordinate>>
          valueBearingCoordinatesBySheet,
      String sheetName) {
    Set<XlsxRoundTripExpectedStateSupport.CellCoordinate> coordinates =
        valueBearingCoordinatesBySheet.get(sheetName);
    if (coordinates == null || coordinates.isEmpty()) {
      return 0;
    }
    int lastRowIndex = -1;
    for (XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate : coordinates) {
      if (coordinate.rowIndex() > lastRowIndex) {
        lastRowIndex = coordinate.rowIndex();
      }
    }
    return lastRowIndex + 1;
  }

  static boolean isValueBearing(dev.erst.gridgrind.excel.ExcelCellValue value) {
    return switch (value) {
      case dev.erst.gridgrind.excel.ExcelCellValue.BlankValue _ -> false;
      case dev.erst.gridgrind.excel.ExcelCellValue.TextValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.RichTextValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.NumberValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.BooleanValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.DateValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.DateTimeValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.FormulaValue _ -> true;
    };
  }

  static List<XlsxRoundTripExpectedStateSupport.CellCoordinate> sortedCoordinates(
      Set<XlsxRoundTripExpectedStateSupport.CellCoordinate> coordinates) {
    java.util.ArrayList<XlsxRoundTripExpectedStateSupport.CellCoordinate> sorted =
        new java.util.ArrayList<>(coordinates);
    sorted.sort(
        (left, right) -> {
          int rowComparison = Integer.compare(left.rowIndex(), right.rowIndex());
          return rowComparison != 0
              ? rowComparison
              : Integer.compare(left.columnIndex(), right.columnIndex());
        });
    return List.copyOf(sorted);
  }

  static void forEachCell(
      String range, Consumer<XlsxRoundTripExpectedStateSupport.CellCoordinate> consumer) {
    CellRangeAddress cellRange = CellRangeAddress.valueOf(range);
    for (int rowIndex = cellRange.getFirstRow(); rowIndex <= cellRange.getLastRow(); rowIndex++) {
      for (int columnIndex = cellRange.getFirstColumn();
          columnIndex <= cellRange.getLastColumn();
          columnIndex++) {
        consumer.accept(
            new XlsxRoundTripExpectedStateSupport.CellCoordinate(rowIndex, columnIndex));
      }
    }
  }
}
