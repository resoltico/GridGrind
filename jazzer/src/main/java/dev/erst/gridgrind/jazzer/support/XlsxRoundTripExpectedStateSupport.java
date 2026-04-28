package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelChartSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelDataValidationSnapshot;
import dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPivotTableSelection;
import dev.erst.gridgrind.excel.ExcelPivotTableSnapshot;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelRichTextSnapshot;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.WorkbookReadResult;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.Consumer;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

/** Captures the pre-save workbook state that must survive `.xlsx` round-trips. */
final class XlsxRoundTripExpectedStateSupport {
  private XlsxRoundTripExpectedStateSupport() {}

  static ExpectedWorkbookState expectedWorkbookState(
      ExcelWorkbook workbook, List<WorkbookCommand> commands) throws IOException {
    ExpectedWorkbookFootprint footprint = expectedWorkbookFootprint(commands);
    Map<String, List<ExcelCellSnapshot>> candidateSnapshots =
        expectedCellSnapshots(workbook, footprint);
    return new ExpectedWorkbookState(
        expectedStyles(candidateSnapshots, XlsxRoundTripVerifier.defaultStyleSnapshot()),
        expectedMetadata(candidateSnapshots),
        expectedRichText(candidateSnapshots),
        expectedNamedRanges(workbook),
        expectedWorkbookSummary(workbook),
        expectedSheetSummaries(workbook),
        expectedSheetLayouts(workbook),
        expectedDataValidations(workbook),
        expectedConditionalFormatting(workbook),
        expectedAutofilters(workbook),
        expectedDrawingObjects(workbook),
        expectedCharts(workbook),
        expectedPivots(workbook),
        expectedTables(workbook));
  }

  static Map<String, ExpectedSheetLayoutState> expectedSheetLayouts(ExcelWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");

    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    LinkedHashMap<String, ExpectedSheetLayoutState> expectedLayouts = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      WorkbookReadResult.SheetLayoutResult layoutResult =
          (WorkbookReadResult.SheetLayoutResult)
              readExecutor
                  .apply(workbook, new WorkbookReadCommand.GetSheetLayout("layout", sheetName))
                  .getFirst();
      expectedLayouts.put(sheetName, expectedSheetLayout(layoutResult.layout()));
    }
    return Map.copyOf(expectedLayouts);
  }

  static ExpectedSheetLayoutState expectedSheetLayout(WorkbookReadResult.SheetLayout layout) {
    Objects.requireNonNull(layout, "layout must not be null");

    LinkedHashMap<Integer, ExpectedRowLayoutState> expectedRows = new LinkedHashMap<>();
    for (WorkbookReadResult.RowLayout row : layout.rows()) {
      expectedRows.put(
          row.rowIndex(),
          new ExpectedRowLayoutState(row.hidden(), row.outlineLevel(), row.collapsed()));
    }

    LinkedHashMap<Integer, ExpectedColumnLayoutState> expectedColumns = new LinkedHashMap<>();
    for (WorkbookReadResult.ColumnLayout column : layout.columns()) {
      expectedColumns.put(
          column.columnIndex(),
          new ExpectedColumnLayoutState(
              column.hidden(), column.outlineLevel(), column.collapsed()));
    }

    return new ExpectedSheetLayoutState(
        layout.pane(), layout.zoomPercent(), Map.copyOf(expectedRows), Map.copyOf(expectedColumns));
  }

  static ExpectedWorkbookFootprint expectedWorkbookFootprint(List<WorkbookCommand> commands) {
    LinkedHashMap<String, Set<CellCoordinate>> candidateCoordinatesBySheet = new LinkedHashMap<>();
    LinkedHashMap<String, Set<CellCoordinate>> valueBearingCoordinatesBySheet =
        new LinkedHashMap<>();

    for (WorkbookCommand command : commands) {
      switch (command) {
        case WorkbookCommand.CreateSheet createSheet ->
            candidateCoordinatesBySheet.computeIfAbsent(
                createSheet.sheetName(), sheetKey -> new LinkedHashSet<>());
        case WorkbookCommand.RenameSheet renameSheet -> {
          renameExpectedSheetState(
              candidateCoordinatesBySheet, renameSheet.sheetName(), renameSheet.newSheetName());
          renameExpectedSheetState(
              valueBearingCoordinatesBySheet, renameSheet.sheetName(), renameSheet.newSheetName());
        }
        case WorkbookCommand.DeleteSheet deleteSheet -> {
          candidateCoordinatesBySheet.remove(deleteSheet.sheetName());
          valueBearingCoordinatesBySheet.remove(deleteSheet.sheetName());
        }
        case WorkbookCommand.MoveSheet _ -> {
          // Sheet order does not affect candidate or append-location tracking.
        }
        case WorkbookCommand.CopySheet copySheet -> {
          copyExpectedSheetState(
              candidateCoordinatesBySheet, copySheet.sourceSheetName(), copySheet.newSheetName());
          copyExpectedSheetState(
              valueBearingCoordinatesBySheet,
              copySheet.sourceSheetName(),
              copySheet.newSheetName());
        }
        case WorkbookCommand.SetActiveSheet _ -> {
          // Active sheet state does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.SetSelectedSheets _ -> {
          // Selected sheet state does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.SetSheetVisibility _ -> {
          // Visibility state does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.SetSheetProtection _ -> {
          // Sheet protection is tracked independently from cell-level expectations.
        }
        case WorkbookCommand.ClearSheetProtection _ -> {
          // Sheet protection is tracked independently from cell-level expectations.
        }
        case WorkbookCommand.SetWorkbookProtection _ -> {
          // Workbook protection is tracked independently from cell-level expectations.
        }
        case WorkbookCommand.ClearWorkbookProtection _ -> {
          // Workbook protection is tracked independently from cell-level expectations.
        }
        case WorkbookCommand.MergeCells _ -> {
          // Merge state does not add new candidate cells.
        }
        case WorkbookCommand.UnmergeCells _ -> {
          // Unmerge state does not add new candidate cells.
        }
        case WorkbookCommand.SetColumnWidth _ -> {
          // Column width does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.SetRowHeight _ -> {
          // Row height does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.InsertRows _ -> {
          // Structural row state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookCommand.DeleteRows _ -> {
          // Structural row state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookCommand.ShiftRows _ -> {
          // Structural row state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookCommand.InsertColumns _ -> {
          // Structural column state is tracked independently from cell-level round-trip
          // expectations.
        }
        case WorkbookCommand.DeleteColumns _ -> {
          // Structural column state is tracked independently from cell-level round-trip
          // expectations.
        }
        case WorkbookCommand.ShiftColumns _ -> {
          // Structural column state is tracked independently from cell-level round-trip
          // expectations.
        }
        case WorkbookCommand.SetRowVisibility _ -> {
          // Row visibility is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookCommand.SetColumnVisibility _ -> {
          // Column visibility is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookCommand.GroupRows _ -> {
          // Row outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookCommand.UngroupRows _ -> {
          // Row outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookCommand.GroupColumns _ -> {
          // Column outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookCommand.UngroupColumns _ -> {
          // Column outline state is tracked independently from cell-level round-trip expectations.
        }
        case WorkbookCommand.SetSheetPane _ -> {
          // Pane state does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.SetSheetZoom _ -> {
          // Zoom state does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.SetSheetPresentation _ -> {
          // Sheet presentation does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.SetPrintLayout _ -> {
          // Print layout does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.ClearPrintLayout _ -> {
          // Print layout clearing does not affect cell-level round-trip expectations.
        }
        case WorkbookCommand.SetCell setCell -> {
          CellCoordinate coordinate = CellCoordinate.fromAddress(setCell.address());
          recordCandidate(candidateCoordinatesBySheet, setCell.sheetName(), coordinate);
          recordValueBearing(
              valueBearingCoordinatesBySheet,
              setCell.sheetName(),
              coordinate,
              isValueBearing(setCell.value()));
        }
        case WorkbookCommand.SetRange setRange -> {
          CellRangeAddress range = CellRangeAddress.valueOf(setRange.range());
          for (int rowOffset = 0; rowOffset < setRange.rows().size(); rowOffset++) {
            for (int columnOffset = 0;
                columnOffset < setRange.rows().get(rowOffset).size();
                columnOffset++) {
              CellCoordinate coordinate =
                  new CellCoordinate(
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
        case WorkbookCommand.ClearRange clearRange ->
            forEachCell(
                clearRange.range(),
                coordinate -> {
                  recordCandidate(candidateCoordinatesBySheet, clearRange.sheetName(), coordinate);
                  recordValueBearing(
                      valueBearingCoordinatesBySheet, clearRange.sheetName(), coordinate, false);
                });
        case WorkbookCommand.SetArrayFormula setArrayFormula ->
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
        case WorkbookCommand.ClearArrayFormula clearArrayFormula -> {
          CellCoordinate coordinate = CellCoordinate.fromAddress(clearArrayFormula.address());
          recordCandidate(candidateCoordinatesBySheet, clearArrayFormula.sheetName(), coordinate);
          recordValueBearing(
              valueBearingCoordinatesBySheet, clearArrayFormula.sheetName(), coordinate, false);
        }
        case WorkbookCommand.ImportCustomXmlMapping _ -> {
          // Custom-XML imports mutate workbook content outside the fuzz model's candidate-cell set.
        }
        case WorkbookCommand.SetHyperlink setHyperlink ->
            recordCandidate(
                candidateCoordinatesBySheet,
                setHyperlink.sheetName(),
                CellCoordinate.fromAddress(setHyperlink.address()));
        case WorkbookCommand.ClearHyperlink clearHyperlink ->
            recordCandidate(
                candidateCoordinatesBySheet,
                clearHyperlink.sheetName(),
                CellCoordinate.fromAddress(clearHyperlink.address()));
        case WorkbookCommand.SetComment setComment ->
            recordCandidate(
                candidateCoordinatesBySheet,
                setComment.sheetName(),
                CellCoordinate.fromAddress(setComment.address()));
        case WorkbookCommand.ClearComment clearComment ->
            recordCandidate(
                candidateCoordinatesBySheet,
                clearComment.sheetName(),
                CellCoordinate.fromAddress(clearComment.address()));
        case WorkbookCommand.SetPicture _ -> {
          // Drawing objects are validated through workbook readability, not candidate cell grids.
        }
        case WorkbookCommand.SetSignatureLine _ -> {
          // Signature lines are validated through workbook readability, not candidate cell grids.
        }
        case WorkbookCommand.SetShape _ -> {
          // Drawing objects are validated through workbook readability, not candidate cell grids.
        }
        case WorkbookCommand.SetEmbeddedObject _ -> {
          // Drawing objects are validated through workbook readability, not candidate cell grids.
        }
        case WorkbookCommand.SetChart _ -> {
          // Charts are validated through workbook-level drawing and chart snapshots.
        }
        case WorkbookCommand.SetPivotTable _ -> {
          // Pivot tables are validated through workbook-level pivot snapshots.
        }
        case WorkbookCommand.SetDrawingObjectAnchor _ -> {
          // Anchor mutation changes drawing geometry only.
        }
        case WorkbookCommand.DeleteDrawingObject _ -> {
          // Drawing deletion does not alter candidate cell tracking.
        }
        case WorkbookCommand.ApplyStyle applyStyle ->
            forEachCell(
                applyStyle.range(),
                coordinate ->
                    recordCandidate(
                        candidateCoordinatesBySheet, applyStyle.sheetName(), coordinate));
        case WorkbookCommand.SetDataValidation _ -> {
          // Data validations are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.ClearDataValidations _ -> {
          // Data validations are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.SetConditionalFormatting _ -> {
          // Conditional formatting is tracked independently from cell-style and metadata
          // expectations.
        }
        case WorkbookCommand.ClearConditionalFormatting _ -> {
          // Conditional formatting is tracked independently from cell-style and metadata
          // expectations.
        }
        case WorkbookCommand.SetAutofilter _ -> {
          // Autofilters are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.ClearAutofilter _ -> {
          // Autofilters are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.SetTable _ -> {
          // Tables are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.DeleteTable _ -> {
          // Tables are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.DeletePivotTable _ -> {
          // Pivot tables are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.SetNamedRange _ -> {
          // Named ranges are derived directly from the applied workbook state.
        }
        case WorkbookCommand.DeleteNamedRange _ -> {
          // Named ranges are derived directly from the applied workbook state.
        }
        case WorkbookCommand.AppendRow appendRow -> {
          int rowIndex = nextAppendRowIndex(valueBearingCoordinatesBySheet, appendRow.sheetName());
          for (int columnIndex = 0; columnIndex < appendRow.values().size(); columnIndex++) {
            CellCoordinate coordinate = new CellCoordinate(rowIndex, columnIndex);
            recordCandidate(candidateCoordinatesBySheet, appendRow.sheetName(), coordinate);
            recordValueBearing(
                valueBearingCoordinatesBySheet,
                appendRow.sheetName(),
                coordinate,
                isValueBearing(appendRow.values().get(columnIndex)));
          }
        }
        case WorkbookCommand.AutoSizeColumns _ -> {
          // Auto-sizing does not add new candidate cells.
        }
      }
    }

    LinkedHashMap<String, List<CellCoordinate>> snapshotCoordinatesBySheet = new LinkedHashMap<>();
    candidateCoordinatesBySheet.forEach(
        (sheetName, coordinates) -> {
          if (!coordinates.isEmpty()) {
            snapshotCoordinatesBySheet.put(sheetName, sortedCoordinates(coordinates));
          }
        });
    return new ExpectedWorkbookFootprint(Map.copyOf(snapshotCoordinatesBySheet));
  }

  static Map<String, List<ExcelCellSnapshot>> expectedCellSnapshots(
      ExcelWorkbook workbook, ExpectedWorkbookFootprint footprint) {
    LinkedHashMap<String, List<ExcelCellSnapshot>> snapshotsBySheet = new LinkedHashMap<>();
    for (Map.Entry<String, List<CellCoordinate>> entry :
        footprint.candidateCoordinatesBySheet().entrySet()) {
      List<String> addresses = entry.getValue().stream().map(CellCoordinate::a1Address).toList();
      List<ExcelCellSnapshot> snapshots = workbook.sheet(entry.getKey()).snapshotCells(addresses);
      if (!snapshots.isEmpty()) {
        snapshotsBySheet.put(entry.getKey(), snapshots);
      }
    }
    return Map.copyOf(snapshotsBySheet);
  }

  static Map<String, Map<CellCoordinate, ExcelCellStyleSnapshot>> expectedStyles(
      Map<String, List<ExcelCellSnapshot>> candidateSnapshots,
      ExcelCellStyleSnapshot defaultStyle) {
    LinkedHashMap<String, Map<CellCoordinate, ExcelCellStyleSnapshot>> expectedStylesBySheet =
        new LinkedHashMap<>();
    for (Map.Entry<String, List<ExcelCellSnapshot>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<CellCoordinate, ExcelCellStyleSnapshot> expectedStyles = new LinkedHashMap<>();
      for (ExcelCellSnapshot snapshot : entry.getValue()) {
        if (!snapshot.style().equals(defaultStyle)) {
          expectedStyles.put(CellCoordinate.fromAddress(snapshot.address()), snapshot.style());
        }
      }
      if (!expectedStyles.isEmpty()) {
        expectedStylesBySheet.put(entry.getKey(), Map.copyOf(expectedStyles));
      }
    }
    return Map.copyOf(expectedStylesBySheet);
  }

  static Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadata(
      Map<String, List<ExcelCellSnapshot>> candidateSnapshots) {
    LinkedHashMap<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadataBySheet =
        new LinkedHashMap<>();
    for (Map.Entry<String, List<ExcelCellSnapshot>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<CellCoordinate, ExpectedCellMetadata> expectedMetadata = new LinkedHashMap<>();
      for (ExcelCellSnapshot snapshot : entry.getValue()) {
        if (hasMetadata(snapshot.metadata())) {
          expectedMetadata.put(
              CellCoordinate.fromAddress(snapshot.address()),
              ExpectedCellMetadata.from(snapshot.metadata()));
        }
      }
      if (!expectedMetadata.isEmpty()) {
        expectedMetadataBySheet.put(entry.getKey(), Map.copyOf(expectedMetadata));
      }
    }
    return Map.copyOf(expectedMetadataBySheet);
  }

  static Map<String, Map<CellCoordinate, ExcelRichTextSnapshot>> expectedRichText(
      Map<String, List<ExcelCellSnapshot>> candidateSnapshots) {
    LinkedHashMap<String, Map<CellCoordinate, ExcelRichTextSnapshot>> expectedRichTextBySheet =
        new LinkedHashMap<>();
    for (Map.Entry<String, List<ExcelCellSnapshot>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<CellCoordinate, ExcelRichTextSnapshot> expectedRichText = new LinkedHashMap<>();
      for (ExcelCellSnapshot snapshot : entry.getValue()) {
        if (snapshot instanceof ExcelCellSnapshot.TextSnapshot text && text.richText() != null) {
          expectedRichText.put(CellCoordinate.fromAddress(text.address()), text.richText());
        }
      }
      if (!expectedRichText.isEmpty()) {
        expectedRichTextBySheet.put(entry.getKey(), Map.copyOf(expectedRichText));
      }
    }
    return Map.copyOf(expectedRichTextBySheet);
  }

  static Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges(ExcelWorkbook workbook) {
    LinkedHashMap<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges = new LinkedHashMap<>();
    for (ExcelNamedRangeSnapshot namedRange : workbook.namedRanges()) {
      ExcelNamedRangeTarget target =
          switch (namedRange) {
            case ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot -> rangeSnapshot.target();
            case ExcelNamedRangeSnapshot.FormulaSnapshot _ -> null;
          };
      expectedNamedRanges.put(
          new NamedRangeKey(namedRange.name(), namedRange.scope()),
          new ExpectedNamedRange(namedRange.scope(), namedRange.refersToFormula(), target));
    }
    return Map.copyOf(expectedNamedRanges);
  }

  static dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary expectedWorkbookSummary(
      ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    return ((dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult)
            readExecutor
                .apply(workbook, new WorkbookReadCommand.GetWorkbookSummary("workbook-summary"))
                .getFirst())
        .workbook();
  }

  static Map<String, dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary>
      expectedSheetSummaries(ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    LinkedHashMap<String, dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary> expected =
        new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      expected.put(
          sheetName,
          ((dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult)
                  readExecutor
                      .apply(
                          workbook,
                          new WorkbookReadCommand.GetSheetSummary(
                              "sheet-summary-" + sheetName, sheetName))
                      .getFirst())
              .sheet());
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelDataValidationSnapshot>> expectedDataValidations(
      ExcelWorkbook workbook) {
    LinkedHashMap<String, List<ExcelDataValidationSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      List<ExcelDataValidationSnapshot> validations =
          workbook.sheet(sheetName).dataValidations(new ExcelRangeSelection.All());
      if (!validations.isEmpty()) {
        expected.put(sheetName, validations);
      }
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelConditionalFormattingBlockSnapshot>> expectedConditionalFormatting(
      ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    LinkedHashMap<String, List<ExcelConditionalFormattingBlockSnapshot>> expected =
        new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      var result =
          (dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult)
              readExecutor
                  .apply(
                      workbook,
                      new WorkbookReadCommand.GetConditionalFormatting(
                          "conditionalFormatting", sheetName, new ExcelRangeSelection.All()))
                  .getFirst();
      if (!result.conditionalFormattingBlocks().isEmpty()) {
        expected.put(sheetName, result.conditionalFormattingBlocks());
      }
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelAutofilterSnapshot>> expectedAutofilters(ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    LinkedHashMap<String, List<ExcelAutofilterSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      var result =
          (dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult)
              readExecutor
                  .apply(workbook, new WorkbookReadCommand.GetAutofilters("autofilters", sheetName))
                  .getFirst();
      if (!result.autofilters().isEmpty()) {
        expected.put(sheetName, result.autofilters());
      }
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelDrawingObjectSnapshot>> expectedDrawingObjects(
      ExcelWorkbook workbook) {
    LinkedHashMap<String, List<ExcelDrawingObjectSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      expected.put(sheetName, List.copyOf(workbook.sheet(sheetName).drawingObjects()));
    }
    return Map.copyOf(expected);
  }

  static Map<String, List<ExcelChartSnapshot>> expectedCharts(ExcelWorkbook workbook) {
    LinkedHashMap<String, List<ExcelChartSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      expected.put(sheetName, List.copyOf(workbook.sheet(sheetName).charts()));
    }
    return Map.copyOf(expected);
  }

  static List<ExcelPivotTableSnapshot> expectedPivots(ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    var result =
        (dev.erst.gridgrind.excel.WorkbookReadResult.PivotTablesResult)
            readExecutor
                .apply(
                    workbook,
                    new WorkbookReadCommand.GetPivotTables(
                        "pivots", new ExcelPivotTableSelection.All()))
                .getFirst();
    return List.copyOf(result.pivotTables());
  }

  static List<ExcelTableSnapshot> expectedTables(ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    var result =
        (dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult)
            readExecutor
                .apply(
                    workbook,
                    new WorkbookReadCommand.GetTables("tables", new ExcelTableSelection.All()))
                .getFirst();
    return List.copyOf(result.tables());
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
      Map<String, Set<CellCoordinate>> candidateCoordinatesBySheet,
      String sheetName,
      CellCoordinate coordinate) {
    candidateCoordinatesBySheet
        .computeIfAbsent(sheetName, sheetKey -> new LinkedHashSet<>())
        .add(coordinate);
  }

  static void recordValueBearing(
      Map<String, Set<CellCoordinate>> valueBearingCoordinatesBySheet,
      String sheetName,
      CellCoordinate coordinate,
      boolean valueBearing) {
    Set<CellCoordinate> coordinates =
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
      Map<String, Set<CellCoordinate>> valueBearingCoordinatesBySheet, String sheetName) {
    Set<CellCoordinate> coordinates = valueBearingCoordinatesBySheet.get(sheetName);
    if (coordinates == null || coordinates.isEmpty()) {
      return 0;
    }
    int lastRowIndex = -1;
    for (CellCoordinate coordinate : coordinates) {
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

  static List<CellCoordinate> sortedCoordinates(Set<CellCoordinate> coordinates) {
    java.util.ArrayList<CellCoordinate> sorted = new java.util.ArrayList<>(coordinates);
    sorted.sort(
        (left, right) -> {
          int rowComparison = Integer.compare(left.rowIndex(), right.rowIndex());
          return rowComparison != 0
              ? rowComparison
              : Integer.compare(left.columnIndex(), right.columnIndex());
        });
    return List.copyOf(sorted);
  }

  static boolean hasMetadata(ExcelCellMetadataSnapshot metadata) {
    return metadata.hyperlink().isPresent() || metadata.comment().isPresent();
  }

  static void forEachCell(String range, Consumer<CellCoordinate> consumer) {
    CellRangeAddress cellRange = CellRangeAddress.valueOf(range);
    for (int rowIndex = cellRange.getFirstRow(); rowIndex <= cellRange.getLastRow(); rowIndex++) {
      for (int columnIndex = cellRange.getFirstColumn();
          columnIndex <= cellRange.getLastColumn();
          columnIndex++) {
        consumer.accept(new CellCoordinate(rowIndex, columnIndex));
      }
    }
  }

  record ExpectedWorkbookState(
      Map<String, Map<CellCoordinate, ExcelCellStyleSnapshot>> expectedStyles,
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadata,
      Map<String, Map<CellCoordinate, ExcelRichTextSnapshot>> expectedRichText,
      Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges,
      dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary expectedWorkbookSummary,
      Map<String, dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary> expectedSheetSummaries,
      Map<String, ExpectedSheetLayoutState> expectedSheetLayouts,
      Map<String, List<ExcelDataValidationSnapshot>> expectedDataValidations,
      Map<String, List<ExcelConditionalFormattingBlockSnapshot>> expectedConditionalFormatting,
      Map<String, List<ExcelAutofilterSnapshot>> expectedAutofilters,
      Map<String, List<ExcelDrawingObjectSnapshot>> expectedDrawingObjects,
      Map<String, List<ExcelChartSnapshot>> expectedCharts,
      List<ExcelPivotTableSnapshot> expectedPivots,
      List<ExcelTableSnapshot> expectedTables) {}

  record ExpectedWorkbookFootprint(Map<String, List<CellCoordinate>> candidateCoordinatesBySheet) {}

  record CellCoordinate(int rowIndex, int columnIndex) {
    static CellCoordinate fromAddress(String address) {
      CellReference cellReference = new CellReference(address);
      return new CellCoordinate(cellReference.getRow(), cellReference.getCol());
    }

    String a1Address() {
      return new CellReference(rowIndex, columnIndex).formatAsString();
    }
  }

  record ExpectedCellMetadata(ExcelHyperlink hyperlink, ExcelComment comment) {
    static ExpectedCellMetadata from(ExcelCellMetadataSnapshot metadata) {
      return new ExpectedCellMetadata(
          metadata.hyperlink().orElse(null),
          metadata.comment().map(derivedComment -> derivedComment.toPlainComment()).orElse(null));
    }
  }

  record ExpectedSheetLayoutState(
      ExcelSheetPane pane,
      Integer zoomPercent,
      Map<Integer, ExpectedRowLayoutState> rows,
      Map<Integer, ExpectedColumnLayoutState> columns) {}

  record ExpectedRowLayoutState(Boolean hidden, Integer outlineLevel, Boolean collapsed) {}

  record ExpectedColumnLayoutState(Boolean hidden, Integer outlineLevel, Boolean collapsed) {}

  record NamedRangeKey(String name, ExcelNamedRangeScope scope) {
    String displayName() {
      return switch (scope) {
        case ExcelNamedRangeScope.WorkbookScope _ -> "WORKBOOK:" + name;
        case ExcelNamedRangeScope.SheetScope sheetScope ->
            "SHEET:" + sheetScope.sheetName() + ":" + name;
      };
    }
  }

  record ExpectedNamedRange(
      ExcelNamedRangeScope scope, String refersToFormula, ExcelNamedRangeTarget target) {}
}
