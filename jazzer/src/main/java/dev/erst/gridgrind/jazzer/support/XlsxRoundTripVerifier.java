package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSnapshot;
import dev.erst.gridgrind.excel.ExcelCellAlignmentSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellProtectionSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelChartSnapshot;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelDataValidationSnapshot;
import dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelGradientFillSnapshot;
import dev.erst.gridgrind.excel.ExcelGradientStopSnapshot;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelNamedRangeTargets;
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
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelGradientFillGeometry;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelPaneRegion;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PaneType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientStop;

/** Reopens saved workbooks and validates basic `.xlsx` structural invariants. */
public final class XlsxRoundTripVerifier {
  private XlsxRoundTripVerifier() {}

  /** Requires the saved workbook to reopen and preserve bounded structural and style invariants. */
  public static void requireRoundTripReadable(
      ExcelWorkbook workbook, Path workbookPath, List<WorkbookCommand> commands)
      throws IOException {
    if (workbook == null) {
      throw new IllegalArgumentException("workbook must not be null");
    }
    if (workbookPath == null) {
      throw new IllegalArgumentException("workbookPath must not be null");
    }
    if (commands == null) {
      throw new IllegalArgumentException("commands must not be null");
    }
    if (!Files.exists(workbookPath)) {
      throw new IllegalStateException("saved workbook must exist");
    }
    ExpectedWorkbookState expectedWorkbookState = expectedWorkbookState(workbook, commands);
    XlsxCommentPackageInvariantSupport.requireCanonicalCommentPackageState(workbookPath);
    XlsxPicturePackageInvariantSupport.requireCanonicalPicturePackageState(workbookPath);

    try (InputStream inputStream = Files.newInputStream(workbookPath);
        XSSFWorkbook reopenedWorkbook = new XSSFWorkbook(inputStream)) {
      if (reopenedWorkbook.getNumberOfSheets() < 0) {
        throw new IllegalStateException("sheet count must not be negative");
      }
      HashSet<String> names = new HashSet<>();
      for (int sheetIndex = 0; sheetIndex < reopenedWorkbook.getNumberOfSheets(); sheetIndex++) {
        String sheetName = reopenedWorkbook.getSheetName(sheetIndex);
        if (sheetName == null) {
          throw new IllegalStateException("sheetName must not be null");
        }
        if (sheetName.isBlank()) {
          throw new IllegalStateException("sheetName must not be blank");
        }
        if (!names.add(sheetName)) {
          throw new IllegalStateException("sheet names must be unique");
        }
        requireSheetShape(reopenedWorkbook.getSheetAt(sheetIndex));
      }
      requireExpectedStyles(expectedWorkbookState.expectedStyles(), reopenedWorkbook);
      requireExpectedMetadata(expectedWorkbookState.expectedMetadata(), reopenedWorkbook);
      requireExpectedNamedRanges(expectedWorkbookState.expectedNamedRanges(), reopenedWorkbook);
    }
    requireExpectedWorkbookState(expectedWorkbookState, workbookPath);
    requireExpectedSheetLayouts(expectedWorkbookState.expectedSheetLayouts(), workbookPath);
    requireExpectedDataValidations(expectedWorkbookState.expectedDataValidations(), workbookPath);
    requireExpectedConditionalFormatting(
        expectedWorkbookState.expectedConditionalFormatting(), workbookPath);
    requireExpectedAutofilters(expectedWorkbookState.expectedAutofilters(), workbookPath);
    requireExpectedTables(expectedWorkbookState.expectedTables(), workbookPath);
  }

  private static void requireSheetShape(org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    if (sheet.getNumMergedRegions() < 0) {
      throw new IllegalStateException("merged region count must not be negative");
    }
    if (sheet.getPaneInformation() != null
        && (sheet.getPaneInformation().getVerticalSplitPosition() < 0
            || sheet.getPaneInformation().getHorizontalSplitPosition() < 0
            || sheet.getPaneInformation().getVerticalSplitLeftColumn() < 0
            || sheet.getPaneInformation().getHorizontalSplitTopRow() < 0)) {
      throw new IllegalStateException("pane coordinates must not be negative");
    }
    for (Row row : sheet) {
      if (row == null) {
        throw new IllegalStateException("row iterator must not yield null");
      }
      for (Cell cell : row) {
        if (cell == null) {
          throw new IllegalStateException("cell iterator must not yield null");
        }
        requireCellStyleShape(
            (XSSFWorkbook) sheet.getWorkbook(), (XSSFCellStyle) cell.getCellStyle());
      }
    }
  }

  private static void requireCellStyleShape(XSSFWorkbook workbook, XSSFCellStyle style) {
    if (style == null) {
      throw new IllegalStateException("cell style must not be null");
    }
    XSSFFont font = style.getFont();
    if (font == null) {
      throw new IllegalStateException("cell font must not be null");
    }
    if (font.getFontName() == null || font.getFontName().isBlank()) {
      throw new IllegalStateException("font name must not be blank");
    }
    if (font.getFontHeight() <= 0) {
      throw new IllegalStateException("font height must be positive");
    }
    requireColorShape(font.getXSSFColor(), "font color");
    requireColorShape(style.getFillForegroundColorColor(), "fill foreground color");
    requireColorShape(style.getFillBackgroundColorColor(), "fill background color");
    requireColorShape(style.getTopBorderXSSFColor(), "top border color");
    requireColorShape(style.getRightBorderXSSFColor(), "right border color");
    requireColorShape(style.getBottomBorderXSSFColor(), "bottom border color");
    requireColorShape(style.getLeftBorderXSSFColor(), "left border color");
    styleSnapshot(workbook, style);
  }

  private static void requireExpectedStyles(
      Map<String, Map<CellCoordinate, ExcelCellStyleSnapshot>> expectedStyles,
      XSSFWorkbook workbook) {
    if (expectedStyles.isEmpty()) {
      return;
    }
    for (Map.Entry<String, Map<CellCoordinate, ExcelCellStyleSnapshot>> sheetEntry :
        expectedStyles.entrySet()) {
      var sheet = workbook.getSheet(sheetEntry.getKey());
      if (sheet == null) {
        throw new IllegalStateException(
            "expected styled sheet must exist after round-trip: " + sheetEntry.getKey());
      }
      for (Map.Entry<CellCoordinate, ExcelCellStyleSnapshot> cellEntry :
          sheetEntry.getValue().entrySet()) {
        Row row = sheet.getRow(cellEntry.getKey().rowIndex());
        if (row == null) {
          throw new IllegalStateException(
              "expected styled row must exist after round-trip: "
                  + sheetEntry.getKey()
                  + "!"
                  + cellEntry.getKey().a1Address());
        }
        Cell cell = row.getCell(cellEntry.getKey().columnIndex());
        if (cell == null) {
          throw new IllegalStateException(
              "expected styled cell must exist after round-trip: "
                  + sheetEntry.getKey()
                  + "!"
                  + cellEntry.getKey().a1Address());
        }
        requireExpectedStyle(sheetEntry.getKey(), cellEntry.getKey(), cellEntry.getValue(), cell);
      }
    }
  }

  private static void requireExpectedStyle(
      String sheetName,
      CellCoordinate coordinate,
      ExcelCellStyleSnapshot expectedStyle,
      Cell cell) {
    ExcelCellStyleSnapshot actualStyle =
        styleSnapshot(
            (XSSFWorkbook) cell.getSheet().getWorkbook(), (XSSFCellStyle) cell.getCellStyle());
    if (!expectedStyle.equals(actualStyle)) {
      throw new IllegalStateException(
          "style must survive .xlsx round-trip for %s!%s: expected %s but was %s"
              .formatted(sheetName, coordinate.a1Address(), expectedStyle, actualStyle));
    }
  }

  private static void requireExpectedMetadata(
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadata,
      XSSFWorkbook workbook) {
    if (expectedMetadata.isEmpty()) {
      return;
    }
    for (Map.Entry<String, Map<CellCoordinate, ExpectedCellMetadata>> sheetEntry :
        expectedMetadata.entrySet()) {
      var sheet = workbook.getSheet(sheetEntry.getKey());
      if (sheet == null) {
        throw new IllegalStateException(
            "expected metadata sheet must exist after round-trip: " + sheetEntry.getKey());
      }
      for (Map.Entry<CellCoordinate, ExpectedCellMetadata> cellEntry :
          sheetEntry.getValue().entrySet()) {
        Row row = sheet.getRow(cellEntry.getKey().rowIndex());
        if (row == null) {
          throw new IllegalStateException(
              "expected metadata row must exist after round-trip: "
                  + sheetEntry.getKey()
                  + "!"
                  + cellEntry.getKey().a1Address());
        }
        Cell cell = row.getCell(cellEntry.getKey().columnIndex());
        if (cell == null) {
          throw new IllegalStateException(
              "expected metadata cell must exist after round-trip: "
                  + sheetEntry.getKey()
                  + "!"
                  + cellEntry.getKey().a1Address());
        }
        requireExpectedMetadata(
            sheetEntry.getKey(), cellEntry.getKey(), cellEntry.getValue(), cell);
      }
    }
  }

  private static void requireExpectedMetadata(
      String sheetName,
      CellCoordinate coordinate,
      ExpectedCellMetadata expectedMetadata,
      Cell cell) {
    requireEquals(
        sheetName, coordinate, "hyperlink", expectedMetadata.hyperlink(), hyperlink(cell));
    requireEquals(sheetName, coordinate, "comment", expectedMetadata.comment(), comment(cell));
  }

  private static void requireExpectedNamedRanges(
      Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges, XSSFWorkbook workbook) {
    if (expectedNamedRanges.isEmpty()) {
      return;
    }
    Map<NamedRangeKey, ExpectedNamedRange> actualNamedRanges = actualNamedRanges(workbook);
    for (Map.Entry<NamedRangeKey, ExpectedNamedRange> entry : expectedNamedRanges.entrySet()) {
      ExpectedNamedRange actual = actualNamedRanges.get(entry.getKey());
      if (actual == null) {
        throw new IllegalStateException(
            "named range must survive .xlsx round-trip: " + entry.getKey().displayName());
      }
      if (!entry.getValue().equals(actual)) {
        throw new IllegalStateException(
            "named range must survive .xlsx round-trip for "
                + entry.getKey().displayName()
                + ": expected "
                + entry.getValue()
                + " but was "
                + actual);
      }
    }
  }

  private static ExpectedWorkbookState expectedWorkbookState(
      ExcelWorkbook workbook, List<WorkbookCommand> commands) throws IOException {
    ExpectedWorkbookFootprint footprint = expectedWorkbookFootprint(commands);
    Map<String, List<ExcelCellSnapshot>> candidateSnapshots =
        expectedCellSnapshots(workbook, footprint);
    return new ExpectedWorkbookState(
        expectedStyles(candidateSnapshots, defaultStyleSnapshot()),
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

  private static Map<String, ExpectedSheetLayoutState> expectedSheetLayouts(
      ExcelWorkbook workbook) {
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

  private static ExpectedSheetLayoutState expectedSheetLayout(
      WorkbookReadResult.SheetLayout layout) {
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

  private static ExpectedWorkbookFootprint expectedWorkbookFootprint(
      List<WorkbookCommand> commands) {
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

  private static Map<String, List<ExcelCellSnapshot>> expectedCellSnapshots(
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

  private static Map<String, Map<CellCoordinate, ExcelCellStyleSnapshot>> expectedStyles(
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

  private static Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadata(
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

  private static Map<String, Map<CellCoordinate, ExcelRichTextSnapshot>> expectedRichText(
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

  private static Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges(
      ExcelWorkbook workbook) {
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

  private static dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary
      expectedWorkbookSummary(ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    return ((dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult)
            readExecutor
                .apply(workbook, new WorkbookReadCommand.GetWorkbookSummary("workbook-summary"))
                .getFirst())
        .workbook();
  }

  private static Map<String, dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary>
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

  private static Map<String, List<ExcelDataValidationSnapshot>> expectedDataValidations(
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

  private static Map<String, List<ExcelConditionalFormattingBlockSnapshot>>
      expectedConditionalFormatting(ExcelWorkbook workbook) {
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

  private static Map<String, List<ExcelAutofilterSnapshot>> expectedAutofilters(
      ExcelWorkbook workbook) {
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

  private static Map<String, List<ExcelDrawingObjectSnapshot>> expectedDrawingObjects(
      ExcelWorkbook workbook) {
    LinkedHashMap<String, List<ExcelDrawingObjectSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      expected.put(sheetName, List.copyOf(workbook.sheet(sheetName).drawingObjects()));
    }
    return Map.copyOf(expected);
  }

  private static Map<String, List<ExcelChartSnapshot>> expectedCharts(ExcelWorkbook workbook) {
    LinkedHashMap<String, List<ExcelChartSnapshot>> expected = new LinkedHashMap<>();
    for (String sheetName : workbook.sheetNames()) {
      expected.put(sheetName, List.copyOf(workbook.sheet(sheetName).charts()));
    }
    return Map.copyOf(expected);
  }

  private static List<ExcelPivotTableSnapshot> expectedPivots(ExcelWorkbook workbook) {
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

  private static List<ExcelTableSnapshot> expectedTables(ExcelWorkbook workbook) {
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

  private static void requireExpectedDataValidations(
      Map<String, List<ExcelDataValidationSnapshot>> expectedDataValidations, Path workbookPath)
      throws IOException {
    if (expectedDataValidations.isEmpty()) {
      return;
    }
    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      for (Map.Entry<String, List<ExcelDataValidationSnapshot>> entry :
          expectedDataValidations.entrySet()) {
        List<ExcelDataValidationSnapshot> actual =
            workbook.sheet(entry.getKey()).dataValidations(new ExcelRangeSelection.All());
        if (!entry.getValue().equals(actual)) {
          throw new IllegalStateException(
              "data validations changed across round-trip for sheet "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actual);
        }
      }
    }
  }

  private static void requireExpectedWorkbookState(
      ExpectedWorkbookState expectedWorkbookState, Path workbookPath) throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      var actualWorkbookSummary = expectedWorkbookSummary(workbook);
      if (!expectedWorkbookState.expectedWorkbookSummary().equals(actualWorkbookSummary)) {
        throw new IllegalStateException(
            "workbook summary changed across round-trip: expected "
                + expectedWorkbookState.expectedWorkbookSummary()
                + " but was "
                + actualWorkbookSummary);
      }
      for (Map.Entry<String, dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary> entry :
          expectedWorkbookState.expectedSheetSummaries().entrySet()) {
        var actualSheetSummary =
            ((dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult)
                    new WorkbookReadExecutor()
                        .apply(
                            workbook,
                            new WorkbookReadCommand.GetSheetSummary(
                                "sheet-summary-" + entry.getKey(), entry.getKey()))
                        .getFirst())
                .sheet();
        if (!entry.getValue().equals(actualSheetSummary)) {
          throw new IllegalStateException(
              "sheet summary changed across round-trip for "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actualSheetSummary);
        }
      }
      for (Map.Entry<String, Map<CellCoordinate, ExcelRichTextSnapshot>> entry :
          expectedWorkbookState.expectedRichText().entrySet()) {
        for (Map.Entry<CellCoordinate, ExcelRichTextSnapshot> cellEntry :
            entry.getValue().entrySet()) {
          ExcelCellSnapshot actualSnapshot =
              workbook.sheet(entry.getKey()).snapshotCell(cellEntry.getKey().a1Address());
          if (!(actualSnapshot instanceof ExcelCellSnapshot.TextSnapshot textSnapshot)) {
            throw new IllegalStateException(
                "rich text cell must reopen as STRING for "
                    + entry.getKey()
                    + "!"
                    + cellEntry.getKey().a1Address());
          }
          requireEquals(
              entry.getKey(),
              cellEntry.getKey(),
              "stringValue",
              cellEntry.getValue().plainText(),
              textSnapshot.stringValue());
          requireEquals(
              entry.getKey(),
              cellEntry.getKey(),
              "richText",
              cellEntry.getValue(),
              textSnapshot.richText());
        }
      }
      for (Map.Entry<String, List<ExcelDrawingObjectSnapshot>> entry :
          expectedWorkbookState.expectedDrawingObjects().entrySet()) {
        List<ExcelDrawingObjectSnapshot> actualDrawingObjects =
            workbook.sheet(entry.getKey()).drawingObjects();
        if (!entry.getValue().equals(actualDrawingObjects)) {
          throw new IllegalStateException(
              "drawing objects changed across round-trip for "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actualDrawingObjects);
        }
      }
      for (Map.Entry<String, List<ExcelChartSnapshot>> entry :
          expectedWorkbookState.expectedCharts().entrySet()) {
        List<ExcelChartSnapshot> actualCharts = workbook.sheet(entry.getKey()).charts();
        if (!entry.getValue().equals(actualCharts)) {
          throw new IllegalStateException(
              "charts changed across round-trip for "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actualCharts);
        }
      }
      List<ExcelPivotTableSnapshot> actualPivots =
          ((dev.erst.gridgrind.excel.WorkbookReadResult.PivotTablesResult)
                  new WorkbookReadExecutor()
                      .apply(
                          workbook,
                          new WorkbookReadCommand.GetPivotTables(
                              "pivots", new ExcelPivotTableSelection.All()))
                      .getFirst())
              .pivotTables();
      if (!expectedWorkbookState.expectedPivots().equals(actualPivots)) {
        throw new IllegalStateException(
            "pivot tables changed across round-trip: expected "
                + expectedWorkbookState.expectedPivots()
                + " but was "
                + actualPivots);
      }
    }
  }

  private static void requireExpectedSheetLayouts(
      Map<String, ExpectedSheetLayoutState> expectedSheetLayouts, Path workbookPath)
      throws IOException {
    if (expectedSheetLayouts.isEmpty()) {
      return;
    }
    try (InputStream inputStream = Files.newInputStream(workbookPath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
      for (Map.Entry<String, ExpectedSheetLayoutState> entry : expectedSheetLayouts.entrySet()) {
        var sheet = workbook.getSheet(entry.getKey());
        if (sheet == null) {
          throw new IllegalStateException(
              "expected sheet layout sheet must exist after round-trip: " + entry.getKey());
        }
        requireExpectedSheetLayout(entry.getKey(), entry.getValue(), sheet);
      }
    }
  }

  private static void requireExpectedSheetLayout(
      String sheetName,
      ExpectedSheetLayoutState expectedSheetLayout,
      org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    ExcelSheetPane actualPane = pane(sheet);
    int actualZoomPercent = zoomPercent(sheet);
    if (expectedSheetLayout.pane() != null && !expectedSheetLayout.pane().equals(actualPane)) {
      throw new IllegalStateException(
          "sheet pane changed across round-trip for "
              + sheetName
              + ": expected "
              + expectedSheetLayout.pane()
              + " but was "
              + actualPane);
    }
    if (expectedSheetLayout.zoomPercent() != null
        && !expectedSheetLayout.zoomPercent().equals(actualZoomPercent)) {
      throw new IllegalStateException(
          "sheet zoom changed across round-trip for "
              + sheetName
              + ": expected "
              + expectedSheetLayout.zoomPercent()
              + " but was "
              + actualZoomPercent);
    }
    for (Map.Entry<Integer, ExpectedRowLayoutState> rowEntry :
        expectedSheetLayout.rows().entrySet()) {
      requireExpectedRowLayout(sheetName, rowEntry.getKey(), rowEntry.getValue(), sheet);
    }
    for (Map.Entry<Integer, ExpectedColumnLayoutState> columnEntry :
        expectedSheetLayout.columns().entrySet()) {
      requireExpectedColumnLayout(sheetName, columnEntry.getKey(), columnEntry.getValue(), sheet);
    }
  }

  private static void requireExpectedRowLayout(
      String sheetName,
      int rowIndex,
      ExpectedRowLayoutState expectedRowLayout,
      org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    Row row = sheet.getRow(rowIndex);
    Boolean hidden =
        row instanceof org.apache.poi.xssf.usermodel.XSSFRow xssfRow && xssfRow.getZeroHeight();
    Integer outlineLevel = row == null ? 0 : Math.max(0, row.getOutlineLevel());
    Boolean collapsed =
        row instanceof org.apache.poi.xssf.usermodel.XSSFRow xssfRow
            && xssfRow.getCTRow().getCollapsed();
    requireLayoutField(sheetName, "row", rowIndex, "hidden", expectedRowLayout.hidden(), hidden);
    requireLayoutField(
        sheetName, "row", rowIndex, "outlineLevel", expectedRowLayout.outlineLevel(), outlineLevel);
    requireLayoutField(
        sheetName, "row", rowIndex, "collapsed", expectedRowLayout.collapsed(), collapsed);
  }

  private static void requireExpectedColumnLayout(
      String sheetName,
      int columnIndex,
      ExpectedColumnLayoutState expectedColumnLayout,
      org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    Boolean hidden = sheet.isColumnHidden(columnIndex);
    Integer outlineLevel = sheet.getColumnOutlineLevel(columnIndex);
    Boolean collapsed = columnCollapsed(sheet, columnIndex);
    requireLayoutField(
        sheetName, "column", columnIndex, "hidden", expectedColumnLayout.hidden(), hidden);
    requireLayoutField(
        sheetName,
        "column",
        columnIndex,
        "outlineLevel",
        expectedColumnLayout.outlineLevel(),
        outlineLevel);
    requireLayoutField(
        sheetName, "column", columnIndex, "collapsed", expectedColumnLayout.collapsed(), collapsed);
  }

  private static void requireExpectedConditionalFormatting(
      Map<String, List<ExcelConditionalFormattingBlockSnapshot>> expectedConditionalFormatting,
      Path workbookPath)
      throws IOException {
    if (expectedConditionalFormatting.isEmpty()) {
      return;
    }
    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
      for (Map.Entry<String, List<ExcelConditionalFormattingBlockSnapshot>> entry :
          expectedConditionalFormatting.entrySet()) {
        var actual =
            ((dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult)
                    readExecutor
                        .apply(
                            workbook,
                            new WorkbookReadCommand.GetConditionalFormatting(
                                "conditionalFormatting",
                                entry.getKey(),
                                new ExcelRangeSelection.All()))
                        .getFirst())
                .conditionalFormattingBlocks();
        if (!entry.getValue().equals(actual)) {
          throw new IllegalStateException(
              "conditional formatting changed across round-trip for sheet "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actual);
        }
      }
    }
  }

  private static void requireExpectedAutofilters(
      Map<String, List<ExcelAutofilterSnapshot>> expectedAutofilters, Path workbookPath)
      throws IOException {
    if (expectedAutofilters.isEmpty()) {
      return;
    }
    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
      for (Map.Entry<String, List<ExcelAutofilterSnapshot>> entry :
          expectedAutofilters.entrySet()) {
        var actual =
            ((dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult)
                    readExecutor
                        .apply(
                            workbook,
                            new WorkbookReadCommand.GetAutofilters("autofilters", entry.getKey()))
                        .getFirst())
                .autofilters();
        if (!entry.getValue().equals(actual)) {
          throw new IllegalStateException(
              "autofilters changed across round-trip for sheet "
                  + entry.getKey()
                  + ": expected "
                  + entry.getValue()
                  + " but was "
                  + actual);
        }
      }
    }
  }

  private static void requireExpectedTables(
      List<ExcelTableSnapshot> expectedTables, Path workbookPath) throws IOException {
    if (expectedTables.isEmpty()) {
      return;
    }
    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
      var actual =
          ((dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult)
                  readExecutor
                      .apply(
                          workbook,
                          new WorkbookReadCommand.GetTables(
                              "tables", new ExcelTableSelection.All()))
                      .getFirst())
              .tables();
      if (!expectedTables.equals(actual)) {
        throw new IllegalStateException(
            "tables changed across round-trip: expected " + expectedTables + " but was " + actual);
      }
    }
  }

  private static void requireLayoutField(
      String sheetName, String axis, int index, String fieldName, Object expected, Object actual) {
    if (expected == null || expected.equals(actual)) {
      return;
    }
    throw new IllegalStateException(
        "sheet layout %s %d field %s changed across round-trip for %s: expected %s but was %s"
            .formatted(axis, index, fieldName, sheetName, expected, actual));
  }

  private static ExcelSheetPane pane(org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    PaneInformation paneInformation = sheet.getPaneInformation();
    if (paneInformation == null) {
      return new ExcelSheetPane.None();
    }
    if (paneInformation.isFreezePane()) {
      return new ExcelSheetPane.Frozen(
          paneInformation.getVerticalSplitPosition(),
          paneInformation.getHorizontalSplitPosition(),
          paneInformation.getVerticalSplitLeftColumn(),
          paneInformation.getHorizontalSplitTopRow());
    }
    return new ExcelSheetPane.Split(
        paneInformation.getVerticalSplitPosition(),
        paneInformation.getHorizontalSplitPosition(),
        paneInformation.getVerticalSplitLeftColumn(),
        paneInformation.getHorizontalSplitTopRow(),
        paneRegion(paneInformation.getActivePaneType()));
  }

  private static int zoomPercent(org.apache.poi.xssf.usermodel.XSSFSheet sheet) {
    var sheetView = sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0);
    return sheetView.isSetZoomScale() ? Math.toIntExact(sheetView.getZoomScale()) : 100;
  }

  private static ExcelPaneRegion paneRegion(PaneType paneType) {
    return switch (paneType) {
      case UPPER_LEFT -> ExcelPaneRegion.UPPER_LEFT;
      case UPPER_RIGHT -> ExcelPaneRegion.UPPER_RIGHT;
      case LOWER_LEFT -> ExcelPaneRegion.LOWER_LEFT;
      case LOWER_RIGHT -> ExcelPaneRegion.LOWER_RIGHT;
    };
  }

  private static boolean columnCollapsed(
      org.apache.poi.xssf.usermodel.XSSFSheet sheet, int columnIndex) {
    for (CTCols cols : sheet.getCTWorksheet().getColsList()) {
      for (CTCol col : cols.getColList()) {
        if (columnIndex + 1 >= col.getMin() && columnIndex + 1 <= col.getMax()) {
          return col.getCollapsed();
        }
      }
    }
    return false;
  }

  private static ExcelCellStyleSnapshot defaultStyleSnapshot() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Default");
      return workbook.sheet("Default").snapshotCell("A1").style();
    }
  }

  private static <T> void renameExpectedSheetState(
      Map<String, T> valuesBySheet, String sheetName, String newSheetName) {
    T values = valuesBySheet.remove(sheetName);
    if (values != null) {
      valuesBySheet.put(newSheetName, values);
    }
  }

  private static <T> void copyExpectedSheetState(
      Map<String, T> valuesBySheet, String sourceSheetName, String newSheetName) {
    T values = valuesBySheet.get(sourceSheetName);
    if (values != null) {
      valuesBySheet.put(newSheetName, values);
    }
  }

  private static void recordCandidate(
      Map<String, Set<CellCoordinate>> candidateCoordinatesBySheet,
      String sheetName,
      CellCoordinate coordinate) {
    candidateCoordinatesBySheet
        .computeIfAbsent(sheetName, sheetKey -> new LinkedHashSet<>())
        .add(coordinate);
  }

  private static void recordValueBearing(
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

  private static int nextAppendRowIndex(
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

  private static boolean isValueBearing(dev.erst.gridgrind.excel.ExcelCellValue value) {
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

  private static List<CellCoordinate> sortedCoordinates(Set<CellCoordinate> coordinates) {
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

  private static boolean hasMetadata(ExcelCellMetadataSnapshot metadata) {
    return metadata.hyperlink().isPresent() || metadata.comment().isPresent();
  }

  private static void forEachCell(
      String range, java.util.function.Consumer<CellCoordinate> consumer) {
    CellRangeAddress cellRange = CellRangeAddress.valueOf(range);
    for (int rowIndex = cellRange.getFirstRow(); rowIndex <= cellRange.getLastRow(); rowIndex++) {
      for (int columnIndex = cellRange.getFirstColumn();
          columnIndex <= cellRange.getLastColumn();
          columnIndex++) {
        consumer.accept(new CellCoordinate(rowIndex, columnIndex));
      }
    }
  }

  private static String resolveNumberFormat(String numberFormat) {
    return numberFormat == null || numberFormat.isBlank() ? "General" : numberFormat;
  }

  private static ExcelCellStyleSnapshot styleSnapshot(XSSFWorkbook workbook, XSSFCellStyle style) {
    XSSFFont font = style.getFont();
    return new ExcelCellStyleSnapshot(
        resolveNumberFormat(style.getDataFormatString()),
        new ExcelCellAlignmentSnapshot(
            style.getWrapText(),
            fromPoi(style.getAlignment()),
            fromPoi(style.getVerticalAlignment()),
            style.getRotation(),
            style.getIndention()),
        new ExcelCellFontSnapshot(
            font.getBold(),
            font.getItalic(),
            font.getFontName(),
            new ExcelFontHeight(font.getFontHeight()),
            toColorSnapshot(font.getXSSFColor()),
            font.getUnderline() != org.apache.poi.ss.usermodel.Font.U_NONE,
            font.getStrikeout()),
        fillSnapshot(workbook, style),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(
                fromPoi(style.getBorderTop()), toColorSnapshot(style.getTopBorderXSSFColor())),
            new ExcelBorderSideSnapshot(
                fromPoi(style.getBorderRight()), toColorSnapshot(style.getRightBorderXSSFColor())),
            new ExcelBorderSideSnapshot(
                fromPoi(style.getBorderBottom()),
                toColorSnapshot(style.getBottomBorderXSSFColor())),
            new ExcelBorderSideSnapshot(
                fromPoi(style.getBorderLeft()), toColorSnapshot(style.getLeftBorderXSSFColor()))),
        new ExcelCellProtectionSnapshot(style.getLocked(), style.getHidden()));
  }

  private static ExcelCellFillSnapshot fillSnapshot(XSSFWorkbook workbook, XSSFCellStyle style) {
    XSSFCellFill fill = workbook.getStylesSource().getFillAt((int) style.getCoreXf().getFillId());
    if (fill.getCTFill().isSetGradientFill()) {
      return new ExcelCellFillSnapshot(
          ExcelFillPattern.NONE,
          null,
          null,
          gradientFillSnapshot(workbook, fill.getCTFill().getGradientFill()));
    }
    ExcelFillPattern pattern = fromPoi(style.getFillPattern());
    if (pattern == ExcelFillPattern.NONE) {
      return new ExcelCellFillSnapshot(pattern, null, null);
    }
    return new ExcelCellFillSnapshot(
        pattern,
        toColorSnapshot(style.getFillForegroundColorColor()),
        pattern == ExcelFillPattern.SOLID
            ? null
            : toColorSnapshot(style.getFillBackgroundColorColor()));
  }

  private static ExcelGradientFillSnapshot gradientFillSnapshot(
      XSSFWorkbook workbook, CTGradientFill fill) {
    Double left = fill.isSetLeft() ? fill.getLeft() : null;
    Double right = fill.isSetRight() ? fill.getRight() : null;
    Double top = fill.isSetTop() ? fill.getTop() : null;
    Double bottom = fill.isSetBottom() ? fill.getBottom() : null;
    List<ExcelGradientStopSnapshot> stops =
        java.util.Arrays.stream(fill.getStopArray())
            .map(stop -> gradientStopSnapshot(workbook, stop))
            .toList();
    return new ExcelGradientFillSnapshot(
        ExcelGradientFillGeometry.effectiveType(
            fill.isSetType() ? fill.getType().toString() : null, left, right, top, bottom),
        fill.isSetDegree() ? fill.getDegree() : null,
        left,
        right,
        top,
        bottom,
        stops);
  }

  private static ExcelGradientStopSnapshot gradientStopSnapshot(
      XSSFWorkbook workbook, CTGradientStop stop) {
    XSSFColor color =
        XSSFColor.from(stop.getColor(), workbook.getStylesSource().getIndexedColors());
    ThemesTable themes = workbook.getStylesSource().getTheme();
    if (themes != null) {
      themes.inheritFromThemeAsRequired(color);
    }
    return new ExcelGradientStopSnapshot(stop.getPosition(), toColorSnapshot(color));
  }

  private static void requireColorShape(XSSFColor color, String label) {
    if (color == null || color.getRGB() == null) {
      return;
    }
    if (color.getRGB().length != 3) {
      throw new IllegalStateException(label + " must resolve to a 3-byte RGB value");
    }
  }

  private static ExcelColorSnapshot toColorSnapshot(XSSFColor color) {
    if (color == null) {
      return null;
    }
    byte[] rgb = color.getRGB();
    String rgbHex = null;
    if (rgb != null) {
      if (rgb.length != 3) {
        return null;
      }
      rgbHex = "#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }
    Integer theme = color.isThemed() ? color.getTheme() : null;
    Integer indexed = color.isIndexed() ? Short.toUnsignedInt(color.getIndexed()) : null;
    Double tint = color.hasTint() ? color.getTint() : null;
    if (theme != null || indexed != null) {
      rgbHex = null;
    }
    if (rgbHex == null && theme == null && indexed == null) {
      return null;
    }
    return new ExcelColorSnapshot(rgbHex, theme, indexed, tint);
  }

  private static ExcelHorizontalAlignment fromPoi(HorizontalAlignment alignment) {
    return ExcelHorizontalAlignment.valueOf(alignment.name());
  }

  private static ExcelVerticalAlignment fromPoi(VerticalAlignment alignment) {
    return ExcelVerticalAlignment.valueOf(alignment.name());
  }

  private static ExcelBorderStyle fromPoi(BorderStyle borderStyle) {
    return ExcelBorderStyle.valueOf(borderStyle.name());
  }

  private static ExcelFillPattern fromPoi(FillPatternType pattern) {
    return switch (pattern) {
      case NO_FILL -> ExcelFillPattern.NONE;
      case SOLID_FOREGROUND -> ExcelFillPattern.SOLID;
      case FINE_DOTS -> ExcelFillPattern.FINE_DOTS;
      case ALT_BARS -> ExcelFillPattern.ALT_BARS;
      case SPARSE_DOTS -> ExcelFillPattern.SPARSE_DOTS;
      case THICK_HORZ_BANDS -> ExcelFillPattern.THICK_HORIZONTAL_BANDS;
      case THICK_VERT_BANDS -> ExcelFillPattern.THICK_VERTICAL_BANDS;
      case THICK_BACKWARD_DIAG -> ExcelFillPattern.THICK_BACKWARD_DIAGONAL;
      case THICK_FORWARD_DIAG -> ExcelFillPattern.THICK_FORWARD_DIAGONAL;
      case BIG_SPOTS -> ExcelFillPattern.BIG_SPOTS;
      case BRICKS -> ExcelFillPattern.BRICKS;
      case THIN_HORZ_BANDS -> ExcelFillPattern.THIN_HORIZONTAL_BANDS;
      case THIN_VERT_BANDS -> ExcelFillPattern.THIN_VERTICAL_BANDS;
      case THIN_BACKWARD_DIAG -> ExcelFillPattern.THIN_BACKWARD_DIAGONAL;
      case THIN_FORWARD_DIAG -> ExcelFillPattern.THIN_FORWARD_DIAGONAL;
      case SQUARES -> ExcelFillPattern.SQUARES;
      case DIAMONDS -> ExcelFillPattern.DIAMONDS;
      case LESS_DOTS -> ExcelFillPattern.LESS_DOTS;
      case LEAST_DOTS -> ExcelFillPattern.LEAST_DOTS;
    };
  }

  private static ExcelHyperlink hyperlink(Cell cell) {
    XSSFHyperlink hyperlink = (XSSFHyperlink) cell.getHyperlink();
    if (hyperlink == null || hyperlink.getType() == null) {
      return null;
    }
    String target = hyperlink.getAddress();
    if (target == null || target.isBlank()) {
      return null;
    }
    try {
      return switch (hyperlink.getType()) {
        case URL -> new ExcelHyperlink.Url(target);
        case EMAIL -> new ExcelHyperlink.Email(target);
        case FILE -> new ExcelHyperlink.File(target);
        case DOCUMENT -> new ExcelHyperlink.Document(target);
        case NONE -> null;
      };
    } catch (IllegalArgumentException exception) {
      return null;
    }
  }

  private static ExcelComment comment(Cell cell) {
    var comment = cell.getCellComment();
    if (comment == null || comment.getString() == null) {
      return null;
    }
    String text = comment.getString().getString();
    String author = comment.getAuthor();
    if (text == null || text.isBlank() || author == null || author.isBlank()) {
      return null;
    }
    return new ExcelComment(text, author, comment.isVisible());
  }

  private static Map<NamedRangeKey, ExpectedNamedRange> actualNamedRanges(XSSFWorkbook workbook) {
    LinkedHashMap<NamedRangeKey, ExpectedNamedRange> actualNamedRanges = new LinkedHashMap<>();
    for (Name name : workbook.getAllNames()) {
      if (!shouldExpose(name)) {
        continue;
      }
      ExcelNamedRangeScope scope = toScope(workbook, name.getSheetIndex());
      String refersToFormula = Objects.requireNonNullElse(name.getRefersToFormula(), "");
      actualNamedRanges.put(
          new NamedRangeKey(name.getNameName(), scope),
          new ExpectedNamedRange(
              scope,
              refersToFormula,
              ExcelNamedRangeTargets.resolveTarget(refersToFormula, scope).orElse(null)));
    }
    return Map.copyOf(actualNamedRanges);
  }

  private static boolean shouldExpose(Name name) {
    String nameName = name.getNameName();
    return !name.isFunctionName()
        && !name.isHidden()
        && nameName != null
        && !nameName.startsWith("_xlnm.")
        && !nameName.startsWith("_XLNM.");
  }

  private static ExcelNamedRangeScope toScope(XSSFWorkbook workbook, int sheetIndex) {
    if (sheetIndex < 0) {
      return new ExcelNamedRangeScope.WorkbookScope();
    }
    return new ExcelNamedRangeScope.SheetScope(workbook.getSheetName(sheetIndex));
  }

  private static void requireEquals(
      String sheetName,
      CellCoordinate coordinate,
      String fieldName,
      Object expected,
      Object actual) {
    if (expected == null) {
      return;
    }
    if (!expected.equals(actual)) {
      throw new IllegalStateException(
          "%s must survive .xlsx round-trip for %s!%s: expected %s but was %s"
              .formatted(fieldName, sheetName, coordinate.a1Address(), expected, actual));
    }
  }

  private record ExpectedWorkbookState(
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

  private record ExpectedWorkbookFootprint(
      Map<String, List<CellCoordinate>> candidateCoordinatesBySheet) {}

  private record CellCoordinate(int rowIndex, int columnIndex) {
    private static CellCoordinate fromAddress(String address) {
      CellReference cellReference = new CellReference(address);
      return new CellCoordinate(cellReference.getRow(), cellReference.getCol());
    }

    private String a1Address() {
      return new CellReference(rowIndex, columnIndex).formatAsString();
    }
  }

  private record ExpectedCellMetadata(ExcelHyperlink hyperlink, ExcelComment comment) {
    private static ExpectedCellMetadata from(ExcelCellMetadataSnapshot metadata) {
      return new ExpectedCellMetadata(
          metadata.hyperlink().orElse(null),
          metadata.comment().map(derivedComment -> derivedComment.toPlainComment()).orElse(null));
    }
  }

  private record ExpectedSheetLayoutState(
      ExcelSheetPane pane,
      Integer zoomPercent,
      Map<Integer, ExpectedRowLayoutState> rows,
      Map<Integer, ExpectedColumnLayoutState> columns) {}

  private record ExpectedRowLayoutState(Boolean hidden, Integer outlineLevel, Boolean collapsed) {}

  private record ExpectedColumnLayoutState(
      Boolean hidden, Integer outlineLevel, Boolean collapsed) {}

  private record NamedRangeKey(String name, ExcelNamedRangeScope scope) {
    private String displayName() {
      return switch (scope) {
        case ExcelNamedRangeScope.WorkbookScope _ -> "WORKBOOK:" + name;
        case ExcelNamedRangeScope.SheetScope sheetScope ->
            "SHEET:" + sheetScope.sheetName() + ":" + name;
      };
    }
  }

  private record ExpectedNamedRange(
      ExcelNamedRangeScope scope, String refersToFormula, ExcelNamedRangeTarget target) {}
}
