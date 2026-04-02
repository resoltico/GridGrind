package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelDataValidationSnapshot;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelHyperlinkType;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelNamedRangeTargets;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Reopens saved workbooks and validates basic `.xlsx` structural invariants. */
public final class XlsxRoundTripVerifier {
  private XlsxRoundTripVerifier() {}

  /** Requires the saved workbook to reopen and preserve bounded structural and style invariants. */
  public static void requireRoundTripReadable(Path workbookPath, List<WorkbookCommand> commands)
      throws IOException {
    if (workbookPath == null) {
      throw new IllegalArgumentException("workbookPath must not be null");
    }
    if (commands == null) {
      throw new IllegalArgumentException("commands must not be null");
    }
    if (!Files.exists(workbookPath)) {
      throw new IllegalStateException("saved workbook must exist");
    }
    ExpectedWorkbookState expectedWorkbookState = expectedWorkbookState(commands);

    try (InputStream inputStream = Files.newInputStream(workbookPath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
      if (workbook.getNumberOfSheets() < 0) {
        throw new IllegalStateException("sheet count must not be negative");
      }
      HashSet<String> names = new HashSet<>();
      for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
        String sheetName = workbook.getSheetName(sheetIndex);
        if (sheetName == null) {
          throw new IllegalStateException("sheetName must not be null");
        }
        if (sheetName.isBlank()) {
          throw new IllegalStateException("sheetName must not be blank");
        }
        if (!names.add(sheetName)) {
          throw new IllegalStateException("sheet names must be unique");
        }
        requireSheetShape(workbook.getSheetAt(sheetIndex));
      }
      requireExpectedStyles(expectedWorkbookState.expectedStyles(), workbook);
      requireExpectedMetadata(expectedWorkbookState.expectedMetadata(), workbook);
      requireExpectedNamedRanges(expectedWorkbookState.expectedNamedRanges(), workbook);
    }
    requireExpectedWorkbookState(expectedWorkbookState, workbookPath);
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
    if (sheet.getPaneInformation() != null && sheet.getPaneInformation().isFreezePane()) {
      if (sheet.getPaneInformation().getVerticalSplitPosition() < 0
          || sheet.getPaneInformation().getHorizontalSplitPosition() < 0
          || sheet.getPaneInformation().getVerticalSplitLeftColumn() < 0
          || sheet.getPaneInformation().getHorizontalSplitTopRow() < 0) {
        throw new IllegalStateException("freeze pane coordinates must not be negative");
      }
    }
    for (Row row : sheet) {
      if (row == null) {
        throw new IllegalStateException("row iterator must not yield null");
      }
      for (Cell cell : row) {
        if (cell == null) {
          throw new IllegalStateException("cell iterator must not yield null");
        }
        requireCellStyleShape((XSSFCellStyle) cell.getCellStyle());
      }
    }
  }

  private static void requireCellStyleShape(XSSFCellStyle style) {
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
    if (style.getFillPattern() == FillPatternType.SOLID_FOREGROUND) {
      XSSFColor color = style.getFillForegroundColorColor();
      if (color != null && color.getRGB() != null && color.getRGB().length != 3) {
        throw new IllegalStateException("solid fill rgb must be 3 bytes when present");
      }
    }
  }

  private static void requireExpectedStyles(
      Map<String, Map<CellCoordinate, ExpectedStyle>> expectedStyles, XSSFWorkbook workbook) {
    if (expectedStyles.isEmpty()) {
      return;
    }
    for (Map.Entry<String, Map<CellCoordinate, ExpectedStyle>> sheetEntry : expectedStyles.entrySet()) {
      var sheet = workbook.getSheet(sheetEntry.getKey());
      if (sheet == null) {
        throw new IllegalStateException("expected styled sheet must exist after round-trip: " + sheetEntry.getKey());
      }
      for (Map.Entry<CellCoordinate, ExpectedStyle> cellEntry : sheetEntry.getValue().entrySet()) {
        Row row = sheet.getRow(cellEntry.getKey().rowIndex());
        if (row == null) {
          throw new IllegalStateException("expected styled row must exist after round-trip: "
              + sheetEntry.getKey() + "!" + cellEntry.getKey().a1Address());
        }
        Cell cell = row.getCell(cellEntry.getKey().columnIndex());
        if (cell == null) {
          throw new IllegalStateException("expected styled cell must exist after round-trip: "
              + sheetEntry.getKey() + "!" + cellEntry.getKey().a1Address());
        }
        requireExpectedStyle(sheetEntry.getKey(), cellEntry.getKey(), cellEntry.getValue(), cell);
      }
    }
  }

  private static void requireExpectedStyle(
      String sheetName, CellCoordinate coordinate, ExpectedStyle expectedStyle, Cell cell) {
    XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();
    XSSFFont font = style.getFont();

    requireEquals(
        sheetName,
        coordinate,
        "numberFormat",
        expectedStyle.numberFormat(),
        resolveNumberFormat(style.getDataFormatString()));
    requireEquals(
        sheetName,
        coordinate,
        "bold",
        expectedStyle.bold(),
        font.getBold());
    requireEquals(
        sheetName,
        coordinate,
        "italic",
        expectedStyle.italic(),
        font.getItalic());
    requireEquals(
        sheetName,
        coordinate,
        "wrapText",
        expectedStyle.wrapText(),
        style.getWrapText());
    requireEquals(
        sheetName,
        coordinate,
        "horizontalAlignment",
        expectedStyle.horizontalAlignment(),
        ExcelHorizontalAlignment.valueOf(style.getAlignment().name()));
    requireEquals(
        sheetName,
        coordinate,
        "verticalAlignment",
        expectedStyle.verticalAlignment(),
        ExcelVerticalAlignment.valueOf(style.getVerticalAlignment().name()));
    requireEquals(
        sheetName,
        coordinate,
        "fontName",
        expectedStyle.fontName(),
        font.getFontName());
    requireEquals(
        sheetName,
        coordinate,
        "fontHeightTwips",
        expectedStyle.fontHeightTwips(),
        Integer.valueOf(font.getFontHeight()));
    requireEquals(
        sheetName,
        coordinate,
        "fontColor",
        expectedStyle.fontColor(),
        toRgbHex(font.getXSSFColor()));
    requireEquals(
        sheetName,
        coordinate,
        "underline",
        expectedStyle.underline(),
        font.getUnderline() != FontUnderline.NONE.getByteValue());
    requireEquals(
        sheetName,
        coordinate,
        "strikeout",
        expectedStyle.strikeout(),
        font.getStrikeout());
    requireEquals(
        sheetName,
        coordinate,
        "fillColor",
        expectedStyle.fillColor(),
        fillColor(style));
    requireEquals(
        sheetName,
        coordinate,
        "borderTop",
        expectedStyle.borderTop(),
        ExcelBorderStyle.valueOf(style.getBorderTop().name()));
    requireEquals(
        sheetName,
        coordinate,
        "borderRight",
        expectedStyle.borderRight(),
        ExcelBorderStyle.valueOf(style.getBorderRight().name()));
    requireEquals(
        sheetName,
        coordinate,
        "borderBottom",
        expectedStyle.borderBottom(),
        ExcelBorderStyle.valueOf(style.getBorderBottom().name()));
    requireEquals(
        sheetName,
        coordinate,
        "borderLeft",
        expectedStyle.borderLeft(),
        ExcelBorderStyle.valueOf(style.getBorderLeft().name()));
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
      for (Map.Entry<CellCoordinate, ExpectedCellMetadata> cellEntry : sheetEntry.getValue().entrySet()) {
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
      String sheetName, CellCoordinate coordinate, ExpectedCellMetadata expectedMetadata, Cell cell) {
    requireEquals(
        sheetName,
        coordinate,
        "hyperlink",
        expectedMetadata.hyperlink(),
        hyperlink(cell));
    requireEquals(
        sheetName,
        coordinate,
        "comment",
        expectedMetadata.comment(),
        comment(cell));
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

  private static ExpectedWorkbookState expectedWorkbookState(List<WorkbookCommand> commands)
      throws IOException {
    ExpectedWorkbookFootprint footprint = expectedWorkbookFootprint(commands);
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      new WorkbookCommandExecutor().apply(workbook, commands);
      Map<String, List<ExcelCellSnapshot>> candidateSnapshots =
          expectedCellSnapshots(workbook, footprint);
      return new ExpectedWorkbookState(
          expectedStyles(candidateSnapshots, defaultStyleSnapshot()),
          expectedMetadata(candidateSnapshots),
          expectedNamedRanges(workbook),
          expectedWorkbookSummary(workbook),
          expectedSheetSummaries(workbook),
          expectedDataValidations(workbook),
          expectedConditionalFormatting(workbook),
          expectedAutofilters(workbook),
          expectedTables(workbook));
    }
  }

  private static ExpectedWorkbookFootprint expectedWorkbookFootprint(List<WorkbookCommand> commands) {
    LinkedHashMap<String, LinkedHashSet<CellCoordinate>> candidateCoordinatesBySheet =
        new LinkedHashMap<>();
    LinkedHashMap<String, LinkedHashSet<CellCoordinate>> valueBearingCoordinatesBySheet =
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
              valueBearingCoordinatesBySheet, copySheet.sourceSheetName(), copySheet.newSheetName());
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
        case WorkbookCommand.FreezePanes _ -> {
          // Freeze panes do not affect cell-level round-trip expectations.
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
        case WorkbookCommand.ApplyStyle applyStyle ->
            forEachCell(
                applyStyle.range(),
                coordinate ->
                    recordCandidate(candidateCoordinatesBySheet, applyStyle.sheetName(), coordinate));
        case WorkbookCommand.SetDataValidation _ -> {
          // Data validations are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.ClearDataValidations _ -> {
          // Data validations are tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.SetConditionalFormatting _ -> {
          // Conditional formatting is tracked independently from cell-style and metadata expectations.
        }
        case WorkbookCommand.ClearConditionalFormatting _ -> {
          // Conditional formatting is tracked independently from cell-style and metadata expectations.
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
        case WorkbookCommand.EvaluateAllFormulas _ -> {
          // Formula evaluation does not add new candidate cells.
        }
        case WorkbookCommand.ForceFormulaRecalculationOnOpen _ -> {
          // Force-recalc flags do not add new candidate cells.
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
      List<String> addresses =
          entry.getValue().stream().map(CellCoordinate::a1Address).toList();
      List<ExcelCellSnapshot> snapshots = workbook.sheet(entry.getKey()).snapshotCells(addresses);
      if (!snapshots.isEmpty()) {
        snapshotsBySheet.put(entry.getKey(), snapshots);
      }
    }
    return Map.copyOf(snapshotsBySheet);
  }

  private static Map<String, Map<CellCoordinate, ExpectedStyle>> expectedStyles(
      Map<String, List<ExcelCellSnapshot>> candidateSnapshots, ExcelCellStyleSnapshot defaultStyle) {
    LinkedHashMap<String, Map<CellCoordinate, ExpectedStyle>> expectedStylesBySheet =
        new LinkedHashMap<>();
    for (Map.Entry<String, List<ExcelCellSnapshot>> entry : candidateSnapshots.entrySet()) {
      LinkedHashMap<CellCoordinate, ExpectedStyle> expectedStyles = new LinkedHashMap<>();
      for (ExcelCellSnapshot snapshot : entry.getValue()) {
        if (!snapshot.style().equals(defaultStyle)) {
          expectedStyles.put(
              CellCoordinate.fromAddress(snapshot.address()), ExpectedStyle.from(snapshot.style()));
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

  private static Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges(ExcelWorkbook workbook) {
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

  private static dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary expectedWorkbookSummary(
      ExcelWorkbook workbook) {
    WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
    return ((dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult)
            readExecutor.apply(
                workbook, new WorkbookReadCommand.GetWorkbookSummary("workbook-summary")).getFirst())
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
                  readExecutor.apply(
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
    }
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
      for (Map.Entry<String, List<ExcelAutofilterSnapshot>> entry : expectedAutofilters.entrySet()) {
        var actual =
            ((dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult)
                    readExecutor
                        .apply(
                            workbook,
                            new WorkbookReadCommand.GetAutofilters(
                                "autofilters", entry.getKey()))
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
            "tables changed across round-trip: expected "
                + expectedTables
                + " but was "
                + actual);
      }
    }
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
      Map<String, LinkedHashSet<CellCoordinate>> candidateCoordinatesBySheet,
      String sheetName,
      CellCoordinate coordinate) {
    candidateCoordinatesBySheet
        .computeIfAbsent(sheetName, sheetKey -> new LinkedHashSet<>())
        .add(coordinate);
  }

  private static void recordValueBearing(
      Map<String, LinkedHashSet<CellCoordinate>> valueBearingCoordinatesBySheet,
      String sheetName,
      CellCoordinate coordinate,
      boolean valueBearing) {
    LinkedHashSet<CellCoordinate> coordinates =
        valueBearingCoordinatesBySheet.computeIfAbsent(sheetName, sheetKey -> new LinkedHashSet<>());
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
      Map<String, LinkedHashSet<CellCoordinate>> valueBearingCoordinatesBySheet, String sheetName) {
    LinkedHashSet<CellCoordinate> coordinates = valueBearingCoordinatesBySheet.get(sheetName);
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
      case dev.erst.gridgrind.excel.ExcelCellValue.NumberValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.BooleanValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.DateValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.DateTimeValue _ -> true;
      case dev.erst.gridgrind.excel.ExcelCellValue.FormulaValue _ -> true;
    };
  }

  private static List<CellCoordinate> sortedCoordinates(LinkedHashSet<CellCoordinate> coordinates) {
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

  private static void forEachCell(String range, java.util.function.Consumer<CellCoordinate> consumer) {
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

  private static String fillColor(XSSFCellStyle style) {
    if (style.getFillPattern() != FillPatternType.SOLID_FOREGROUND) {
      return null;
    }
    return toRgbHex(style.getFillForegroundColorColor());
  }

  private static String toRgbHex(XSSFColor color) {
    if (color == null) {
      return null;
    }
    byte[] rgb = color.getRGB();
    if (rgb == null || rgb.length != 3) {
      return null;
    }
    return "#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
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
          "style field %s must survive .xlsx round-trip for %s!%s: expected %s but was %s"
              .formatted(fieldName, sheetName, coordinate.a1Address(), expected, actual));
    }
  }

  private record ExpectedWorkbookState(
      Map<String, Map<CellCoordinate, ExpectedStyle>> expectedStyles,
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadata,
      Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges,
      dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary expectedWorkbookSummary,
      Map<String, dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary> expectedSheetSummaries,
      Map<String, List<ExcelDataValidationSnapshot>> expectedDataValidations,
      Map<String, List<ExcelConditionalFormattingBlockSnapshot>> expectedConditionalFormatting,
      Map<String, List<ExcelAutofilterSnapshot>> expectedAutofilters,
      List<ExcelTableSnapshot> expectedTables) {}

  private record ExpectedWorkbookFootprint(Map<String, List<CellCoordinate>> candidateCoordinatesBySheet) {}

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
          metadata.hyperlink().orElse(null), metadata.comment().orElse(null));
    }
  }

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

  private record ExpectedStyle(
      String numberFormat,
      Boolean bold,
      Boolean italic,
      Boolean wrapText,
      ExcelHorizontalAlignment horizontalAlignment,
      ExcelVerticalAlignment verticalAlignment,
      String fontName,
      Integer fontHeightTwips,
      String fontColor,
      Boolean underline,
      Boolean strikeout,
      String fillColor,
      ExcelBorderStyle borderTop,
      ExcelBorderStyle borderRight,
      ExcelBorderStyle borderBottom,
      ExcelBorderStyle borderLeft) {
    private static ExpectedStyle from(ExcelCellStyleSnapshot style) {
      return new ExpectedStyle(
          style.numberFormat(),
          style.bold(),
          style.italic(),
          style.wrapText(),
          style.horizontalAlignment(),
          style.verticalAlignment(),
          style.fontName(),
          style.fontHeight().twips(),
          style.fontColor(),
          style.underline(),
          style.strikeout(),
          style.fillColor(),
          style.topBorderStyle(),
          style.rightBorderStyle(),
          style.bottomBorderStyle(),
          style.leftBorderStyle());
    }

  }
}
