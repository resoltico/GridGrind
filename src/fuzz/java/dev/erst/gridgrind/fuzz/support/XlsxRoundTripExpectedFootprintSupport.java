package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
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
      List<ResolvedStep> commands) {
    LinkedHashMap<String, Set<XlsxRoundTripExpectedStateSupport.CellCoordinate>>
        candidateCoordinatesBySheet = new LinkedHashMap<>();
    LinkedHashMap<String, Set<XlsxRoundTripExpectedStateSupport.CellCoordinate>>
        valueBearingCoordinatesBySheet = new LinkedHashMap<>();

    for (ResolvedStep command : commands) {
      switch (command.action()) {
        case WorkbookMutationAction.EnsureSheet _ ->
            candidateCoordinatesBySheet.computeIfAbsent(
                ((SheetSelector.ByName) command.target()).name().value(),
                sheetKey -> new LinkedHashSet<>());
        case WorkbookMutationAction.RenameSheet renameSheet -> {
          renameExpectedSheetState(
              candidateCoordinatesBySheet,
              ((SheetSelector.ByName) command.target()).name().value(),
              renameSheet.newSheetName().value());
          renameExpectedSheetState(
              valueBearingCoordinatesBySheet,
              ((SheetSelector.ByName) command.target()).name().value(),
              renameSheet.newSheetName().value());
        }
        case WorkbookMutationAction.DeleteSheet _ -> {
          candidateCoordinatesBySheet.remove(
              ((SheetSelector.ByName) command.target()).name().value());
          valueBearingCoordinatesBySheet.remove(
              ((SheetSelector.ByName) command.target()).name().value());
        }
        case WorkbookMutationAction.MoveSheet _ -> {
          // Sheet order does not affect candidate or append-location tracking.
        }
        case WorkbookMutationAction.CopySheet copySheet -> {
          copyExpectedSheetState(
              candidateCoordinatesBySheet,
              ((SheetSelector.ByName) command.target()).name().value(),
              copySheet.newSheetName().value());
          copyExpectedSheetState(
              valueBearingCoordinatesBySheet,
              ((SheetSelector.ByName) command.target()).name().value(),
              copySheet.newSheetName().value());
        }
        case WorkbookMutationAction.SetActiveSheet _ -> {
          // Active sheet state does not affect cell-level round-trip expectations.
        }
        case WorkbookMutationAction.SetSelectedSheets _ -> {
          // Selected sheet state does not affect cell-level round-trip expectations.
        }
        case WorkbookMutationAction.SetSheetVisibility _ -> {
          // Visibility state does not affect cell-level round-trip expectations.
        }
        case ResolvedMutationAction.SetSheetProtection _ -> {
          // Sheet protection is tracked independently from cell-level expectations.
        }
        case WorkbookMutationAction.ClearSheetProtection _ -> {
          // Sheet protection is tracked independently from cell-level expectations.
        }
        case ResolvedMutationAction.SetWorkbookProtection _ -> {
          // Workbook protection is tracked independently from cell-level expectations.
        }
        case WorkbookMutationAction.ClearWorkbookProtection _ -> {
          // Workbook protection is tracked independently from cell-level expectations.
        }
        case WorkbookMutationAction.MergeCells _ -> {
          // Merge state does not add new candidate cells.
        }
        case WorkbookMutationAction.UnmergeCells _ -> {
          // Unmerge state does not add new candidate cells.
        }
        case WorkbookMutationAction.SetColumnWidth _ -> {
          // Column width does not affect cell-level round-trip expectations.
        }
        case WorkbookMutationAction.SetRowHeight _ -> {
          // Row height does not affect cell-level round-trip expectations.
        }
        case WorkbookMutationAction.InsertRows _ -> {
          // Structural row state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookMutationAction.DeleteRows _ -> {
          // Structural row state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookMutationAction.ShiftRows _ -> {
          // Structural row state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookMutationAction.InsertColumns _ -> {
          // Structural column state is tracked independently from cell-level round-trip
          // expectations.
        }
        case WorkbookMutationAction.DeleteColumns _ -> {
          // Structural column state is tracked independently from cell-level round-trip
          // expectations.
        }
        case WorkbookMutationAction.ShiftColumns _ -> {
          // Structural column state is tracked independently from cell-level round-trip
          // expectations.
        }
        case WorkbookMutationAction.SetRowVisibility _ -> {
          // Row visibility is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookMutationAction.SetColumnVisibility _ -> {
          // Column visibility is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookMutationAction.GroupRows _ -> {
          // Row outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookMutationAction.UngroupRows _ -> {
          // Row outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookMutationAction.GroupColumns _ -> {
          // Column outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookMutationAction.UngroupColumns _ -> {
          // Column outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookMutationAction.SetSheetPane _ -> {
          // Pane state does not affect cell-level round-trip expectations.
        }
        case WorkbookMutationAction.SetSheetZoom _ -> {
          // Zoom state does not affect cell-level round-trip expectations.
        }
        case WorkbookMutationAction.SetSheetPresentation _ -> {
          // Sheet presentation does not affect cell-level round-trip expectations.
        }
        case ResolvedMutationAction.SetPrintLayout _ -> {
          // Print layout does not affect cell-level round-trip expectations.
        }
        case WorkbookMutationAction.ClearPrintLayout _ -> {
          // Print layout clearing does not affect cell-level round-trip expectations.
        }
        case ResolvedMutationAction.SetCell setCell -> {
          XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate =
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                  ((CellSelector.ByAddress) command.target()).address().value());
          recordCandidate(
              candidateCoordinatesBySheet,
              ((CellSelector.ByAddress) command.target()).sheetName().value(),
              coordinate);
          recordValueBearing(
              valueBearingCoordinatesBySheet,
              ((CellSelector.ByAddress) command.target()).sheetName().value(),
              coordinate,
              isValueBearing(setCell.value()));
        }
        case ResolvedMutationAction.SetRange setRange -> {
          CellRangeAddress range =
              CellRangeAddress.valueOf(((RangeSelector.ByRange) command.target()).range().value());
          for (int rowOffset = 0; rowOffset < setRange.rows().size(); rowOffset++) {
            for (int columnOffset = 0;
                columnOffset < setRange.rows().get(rowOffset).size();
                columnOffset++) {
              XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate =
                  new XlsxRoundTripExpectedStateSupport.CellCoordinate(
                      range.getFirstRow() + rowOffset, range.getFirstColumn() + columnOffset);
              recordCandidate(
                  candidateCoordinatesBySheet,
                  ((RangeSelector.ByRange) command.target()).sheetName().value(),
                  coordinate);
              recordValueBearing(
                  valueBearingCoordinatesBySheet,
                  ((RangeSelector.ByRange) command.target()).sheetName().value(),
                  coordinate,
                  isValueBearing(setRange.rows().get(rowOffset).get(columnOffset)));
            }
          }
        }
        case CellMutationAction.ClearRange _ ->
            forEachCell(
                ((RangeSelector.ByRange) command.target()).range().value(),
                coordinate -> {
                  recordCandidate(
                      candidateCoordinatesBySheet,
                      ((RangeSelector.ByRange) command.target()).sheetName().value(),
                      coordinate);
                  recordValueBearing(
                      valueBearingCoordinatesBySheet,
                      ((RangeSelector.ByRange) command.target()).sheetName().value(),
                      coordinate,
                      false);
                });
        case ResolvedMutationAction.SetArrayFormula _ ->
            forEachCell(
                ((RangeSelector.ByRange) command.target()).range().value(),
                coordinate -> {
                  recordCandidate(
                      candidateCoordinatesBySheet,
                      ((RangeSelector.ByRange) command.target()).sheetName().value(),
                      coordinate);
                  recordValueBearing(
                      valueBearingCoordinatesBySheet,
                      ((RangeSelector.ByRange) command.target()).sheetName().value(),
                      coordinate,
                      true);
                });
        case CellMutationAction.ClearArrayFormula _ -> {
          XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate =
              XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                  ((CellSelector.ByAddress) command.target()).address().value());
          recordCandidate(
              candidateCoordinatesBySheet,
              ((CellSelector.ByAddress) command.target()).sheetName().value(),
              coordinate);
          recordValueBearing(
              valueBearingCoordinatesBySheet,
              ((CellSelector.ByAddress) command.target()).sheetName().value(),
              coordinate,
              false);
        }
        case ResolvedMutationAction.ImportCustomXmlMapping _ -> {
          // Custom-XML imports mutate workbook content outside the fuzz model's candidate-cell set.
        }
        case CellMutationAction.SetHyperlink _ ->
            recordCandidate(
                candidateCoordinatesBySheet,
                ((CellSelector.ByAddress) command.target()).sheetName().value(),
                XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                    ((CellSelector.ByAddress) command.target()).address().value()));
        case CellMutationAction.ClearHyperlink _ ->
            recordCandidate(
                candidateCoordinatesBySheet,
                ((CellSelector.ByAddress) command.target()).sheetName().value(),
                XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                    ((CellSelector.ByAddress) command.target()).address().value()));
        case ResolvedMutationAction.SetComment _ ->
            recordCandidate(
                candidateCoordinatesBySheet,
                ((CellSelector.ByAddress) command.target()).sheetName().value(),
                XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                    ((CellSelector.ByAddress) command.target()).address().value()));
        case CellMutationAction.ClearComment _ ->
            recordCandidate(
                candidateCoordinatesBySheet,
                ((CellSelector.ByAddress) command.target()).sheetName().value(),
                XlsxRoundTripExpectedStateSupport.CellCoordinate.fromAddress(
                    ((CellSelector.ByAddress) command.target()).address().value()));
        case ResolvedMutationAction.SetPicture _ -> {
          // Drawing objects are validated through workbook readability, not candidate cell grids.
        }
        case ResolvedMutationAction.SetSignatureLine _ -> {
          // Signature lines are validated through workbook readability, not candidate cell grids.
        }
        case ResolvedMutationAction.SetShape _ -> {
          // Drawing objects are validated through workbook readability, not candidate cell grids.
        }
        case ResolvedMutationAction.SetEmbeddedObject _ -> {
          // Drawing objects are validated through workbook readability, not candidate cell grids.
        }
        case ResolvedMutationAction.SetChart _ -> {
          // Charts are validated through workbook-level drawing and chart snapshots.
        }
        case StructuredMutationAction.SetPivotTable _ -> {
          // Pivot tables are validated through workbook-level pivot snapshots.
        }
        case DrawingMutationAction.SetDrawingObjectAnchor _ -> {
          // Anchor mutation changes drawing geometry only.
        }
        case DrawingMutationAction.DeleteDrawingObject _ -> {
          // Drawing deletion does not alter candidate cell tracking.
        }
        case CellMutationAction.ApplyStyle _ ->
            forEachCell(
                ((RangeSelector.ByRange) command.target()).range().value(),
                coordinate ->
                    recordCandidate(
                        candidateCoordinatesBySheet,
                        ((RangeSelector.ByRange) command.target()).sheetName().value(),
                        coordinate));
        case ResolvedMutationAction.SetDataValidation _ -> {
          // Data validations are tracked independently from cell-style and metadata expectations.
        }
        case StructuredMutationAction.ClearDataValidations _ -> {
          // Data validations are tracked independently from cell-style and metadata expectations.
        }
        case StructuredMutationAction.SetConditionalFormatting _ -> {
          // Conditional formatting is tracked independently from cell-style and metadata
          // expectations.
        }
        case StructuredMutationAction.ClearConditionalFormatting _ -> {
          // Conditional formatting is tracked independently from cell-style and metadata
          // expectations.
        }
        case StructuredMutationAction.SetAutofilter _ -> {
          // Autofilters are tracked independently from cell-style and metadata expectations.
        }
        case StructuredMutationAction.ClearAutofilter _ -> {
          // Autofilters are tracked independently from cell-style and metadata expectations.
        }
        case ResolvedMutationAction.SetTable _ -> {
          // Tables are tracked independently from cell-style and metadata expectations.
        }
        case StructuredMutationAction.DeleteTable _ -> {
          // Tables are tracked independently from cell-style and metadata expectations.
        }
        case StructuredMutationAction.DeletePivotTable _ -> {
          // Pivot tables are tracked independently from cell-style and metadata expectations.
        }
        case StructuredMutationAction.SetNamedRange _ -> {
          // Named ranges are derived directly from the applied workbook state.
        }
        case StructuredMutationAction.DeleteNamedRange _ -> {
          // Named ranges are derived directly from the applied workbook state.
        }
        case ResolvedMutationAction.AppendRow appendRow -> {
          int rowIndex =
              nextAppendRowIndex(
                  valueBearingCoordinatesBySheet,
                  ((SheetSelector.ByName) command.target()).name().value());
          for (int columnIndex = 0; columnIndex < appendRow.values().size(); columnIndex++) {
            XlsxRoundTripExpectedStateSupport.CellCoordinate coordinate =
                new XlsxRoundTripExpectedStateSupport.CellCoordinate(rowIndex, columnIndex);
            recordCandidate(
                candidateCoordinatesBySheet,
                ((SheetSelector.ByName) command.target()).name().value(),
                coordinate);
            recordValueBearing(
                valueBearingCoordinatesBySheet,
                ((SheetSelector.ByName) command.target()).name().value(),
                coordinate,
                isValueBearing(appendRow.values().get(columnIndex)));
          }
        }
        case WorkbookMutationAction.AutoSizeColumns _ -> {
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

  static Map<String, List<CellReport>> expectedCellSnapshots(
      XssfWorkbook workbook,
      XlsxRoundTripExpectedStateSupport.ExpectedWorkbookFootprint footprint) {
    LinkedHashMap<String, List<CellReport>> snapshotsBySheet = new LinkedHashMap<>();
    for (Map.Entry<String, List<XlsxRoundTripExpectedStateSupport.CellCoordinate>> entry :
        footprint.candidateCoordinatesBySheet().entrySet()) {
      List<String> addresses =
          entry.getValue().stream()
              .map(XlsxRoundTripExpectedStateSupport.CellCoordinate::a1Address)
              .toList();
      List<CellReport> snapshots = workbook.sheet(entry.getKey()).cells().snapshotCells(addresses);
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

  static boolean isValueBearing(ResolvedCellWrite value) {
    return switch (value) {
      case CellValue.Blank _ -> false;
      case CellValue.Text _ -> true;
      case ResolvedRichText _ -> true;
      case CellValue.NumberValue _ -> true;
      case CellValue.BooleanValue _ -> true;
      case CellValue.ErrorValue _ -> true;
      case CellValueSpec.Date _ -> true;
      case CellValueSpec.DateTime _ -> true;
      case ResolvedFormula _ -> true;
      case ResolvedRawFormula _ -> true;
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
