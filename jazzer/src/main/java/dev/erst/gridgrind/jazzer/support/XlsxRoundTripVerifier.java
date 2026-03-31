package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelHyperlinkType;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelNamedRangeTargets;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.WorkbookCommand;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.HashSet;
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
    Map<String, Map<CellCoordinate, ExpectedStyle>> expectedStyles = expectedStyles(commands);
    Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadata =
        expectedMetadata(commands);
    Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges = expectedNamedRanges(commands);

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
      requireExpectedStyles(expectedStyles, workbook);
      requireExpectedMetadata(expectedMetadata, workbook);
      requireExpectedNamedRanges(expectedNamedRanges, workbook);
    }
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

  private static Map<String, Map<CellCoordinate, ExpectedStyle>> expectedStyles(
      List<WorkbookCommand> commands) {
    LinkedHashMap<String, Map<CellCoordinate, ExpectedStyle>> expectedStylesBySheet =
        new LinkedHashMap<>();
    for (WorkbookCommand command : commands) {
      switch (command) {
        case WorkbookCommand.CreateSheet createSheet ->
            expectedStylesBySheet.putIfAbsent(createSheet.sheetName(), new HashMap<>());
        case WorkbookCommand.RenameSheet renameSheet ->
            renameExpectedSheet(
                expectedStylesBySheet, renameSheet.sheetName(), renameSheet.newSheetName());
        case WorkbookCommand.DeleteSheet deleteSheet ->
            expectedStylesBySheet.remove(deleteSheet.sheetName());
        case WorkbookCommand.MoveSheet _ -> {
          // Sheet order does not affect cell-level style expectations.
        }
        case WorkbookCommand.MergeCells _ -> {
          // Merge state does not change style persistence expectations.
        }
        case WorkbookCommand.UnmergeCells _ -> {
          // Unmerge state does not change style persistence expectations.
        }
        case WorkbookCommand.SetColumnWidth _ -> {
          // Column width changes do not affect cell styles.
        }
        case WorkbookCommand.SetRowHeight _ -> {
          // Row height changes do not affect cell styles.
        }
        case WorkbookCommand.FreezePanes _ -> {
          // Freeze panes do not affect cell styles.
        }
        case WorkbookCommand.SetCell setCell ->
            clearExpectedRange(expectedStylesBySheet, setCell.sheetName(), setCell.address());
        case WorkbookCommand.SetRange setRange ->
            clearExpectedRange(expectedStylesBySheet, setRange.sheetName(), setRange.range());
        case WorkbookCommand.ClearRange clearRange ->
            clearExpectedRange(expectedStylesBySheet, clearRange.sheetName(), clearRange.range());
        case WorkbookCommand.SetHyperlink _ -> {
          // Hyperlinks are exercised by separate response and persistence invariants.
        }
        case WorkbookCommand.ClearHyperlink _ -> {
          // Hyperlink removal does not affect cell styles.
        }
        case WorkbookCommand.SetComment _ -> {
          // Comments are exercised by separate response and persistence invariants.
        }
        case WorkbookCommand.ClearComment _ -> {
          // Comment removal does not affect cell styles.
        }
        case WorkbookCommand.ApplyStyle applyStyle ->
            applyExpectedStyle(
                expectedStylesBySheet, applyStyle.sheetName(), applyStyle.range(), applyStyle.style());
        case WorkbookCommand.SetNamedRange _ -> {
          // Named ranges are independent of cell-style persistence expectations.
        }
        case WorkbookCommand.DeleteNamedRange _ -> {
          // Named-range deletion is independent of cell-style persistence expectations.
        }
        case WorkbookCommand.AppendRow _ -> {
          // Append operations may reuse style-only rows, but value writes must preserve any
          // existing style state on those cells.
        }
        case WorkbookCommand.AutoSizeColumns _ -> {
          // Auto-sizing does not affect cell styles.
        }
        case WorkbookCommand.EvaluateAllFormulas _ -> {
          // Formula evaluation does not affect cell styles.
        }
        case WorkbookCommand.ForceFormulaRecalculationOnOpen _ -> {
          // Force-recalc flags do not affect cell styles.
        }
      }
    }
    expectedStylesBySheet.entrySet().removeIf(entry -> entry.getValue().isEmpty());
    return Map.copyOf(expectedStylesBySheet);
  }

  private static Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadata(
      List<WorkbookCommand> commands) {
    LinkedHashMap<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadataBySheet =
        new LinkedHashMap<>();
    for (WorkbookCommand command : commands) {
      switch (command) {
        case WorkbookCommand.CreateSheet createSheet ->
            expectedMetadataBySheet.putIfAbsent(createSheet.sheetName(), new HashMap<>());
        case WorkbookCommand.RenameSheet renameSheet ->
            renameExpectedSheet(
                expectedMetadataBySheet, renameSheet.sheetName(), renameSheet.newSheetName());
        case WorkbookCommand.DeleteSheet deleteSheet ->
            expectedMetadataBySheet.remove(deleteSheet.sheetName());
        case WorkbookCommand.MoveSheet _ -> {
          // Sheet order does not affect cell metadata persistence expectations.
        }
        case WorkbookCommand.MergeCells _ -> {
          // Merge state does not affect hyperlink or comment persistence.
        }
        case WorkbookCommand.UnmergeCells _ -> {
          // Unmerge state does not affect hyperlink or comment persistence.
        }
        case WorkbookCommand.SetColumnWidth _ -> {
          // Column width does not affect hyperlink or comment persistence.
        }
        case WorkbookCommand.SetRowHeight _ -> {
          // Row height does not affect hyperlink or comment persistence.
        }
        case WorkbookCommand.FreezePanes _ -> {
          // Freeze panes do not affect hyperlink or comment persistence.
        }
        case WorkbookCommand.SetCell _ -> {
          // Value writes do not implicitly clear hyperlinks or comments.
        }
        case WorkbookCommand.SetRange _ -> {
          // Matrix writes do not implicitly clear hyperlinks or comments.
        }
        case WorkbookCommand.ClearRange clearRange ->
            clearExpectedMetadataRange(
                expectedMetadataBySheet, clearRange.sheetName(), clearRange.range());
        case WorkbookCommand.SetHyperlink setHyperlink ->
            applyExpectedMetadata(
                expectedMetadataBySheet,
                setHyperlink.sheetName(),
                setHyperlink.address(),
                ExpectedCellMetadata.withHyperlink(setHyperlink.target()));
        case WorkbookCommand.ClearHyperlink clearHyperlink ->
            clearExpectedHyperlink(
                expectedMetadataBySheet, clearHyperlink.sheetName(), clearHyperlink.address());
        case WorkbookCommand.SetComment setComment ->
            applyExpectedMetadata(
                expectedMetadataBySheet,
                setComment.sheetName(),
                setComment.address(),
                ExpectedCellMetadata.withComment(setComment.comment()));
        case WorkbookCommand.ClearComment clearComment ->
            clearExpectedComment(
                expectedMetadataBySheet, clearComment.sheetName(), clearComment.address());
        case WorkbookCommand.ApplyStyle _ -> {
          // Styling does not affect hyperlink or comment persistence.
        }
        case WorkbookCommand.SetNamedRange _ -> {
          // Named ranges are tracked independently from cell metadata.
        }
        case WorkbookCommand.DeleteNamedRange _ -> {
          // Named ranges are tracked independently from cell metadata.
        }
        case WorkbookCommand.AppendRow _ -> {
          // Append operations may reuse metadata-only rows, and value writes must preserve any
          // existing hyperlink or comment state on those cells.
        }
        case WorkbookCommand.AutoSizeColumns _ -> {
          // Auto-sizing does not affect hyperlink or comment persistence.
        }
        case WorkbookCommand.EvaluateAllFormulas _ -> {
          // Formula evaluation does not affect hyperlink or comment persistence.
        }
        case WorkbookCommand.ForceFormulaRecalculationOnOpen _ -> {
          // Force-recalc flags do not affect hyperlink or comment persistence.
        }
      }
    }
    expectedMetadataBySheet.entrySet().removeIf(entry -> entry.getValue().isEmpty());
    return Map.copyOf(expectedMetadataBySheet);
  }

  private static Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges(
      List<WorkbookCommand> commands) {
    LinkedHashMap<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges = new LinkedHashMap<>();
    for (WorkbookCommand command : commands) {
      switch (command) {
        case WorkbookCommand.CreateSheet _ -> {
          // Sheet creation alone does not create named ranges.
        }
        case WorkbookCommand.RenameSheet renameSheet ->
            renameExpectedNamedRanges(
                expectedNamedRanges, renameSheet.sheetName(), renameSheet.newSheetName());
        case WorkbookCommand.DeleteSheet deleteSheet ->
            removeExpectedNamedRangesForSheet(expectedNamedRanges, deleteSheet.sheetName());
        case WorkbookCommand.MoveSheet _ -> {
          // Sheet order does not affect named-range persistence.
        }
        case WorkbookCommand.MergeCells _ -> {
          // Merge state does not affect named ranges directly.
        }
        case WorkbookCommand.UnmergeCells _ -> {
          // Unmerge state does not affect named ranges directly.
        }
        case WorkbookCommand.SetColumnWidth _ -> {
          // Column width does not affect named ranges.
        }
        case WorkbookCommand.SetRowHeight _ -> {
          // Row height does not affect named ranges.
        }
        case WorkbookCommand.FreezePanes _ -> {
          // Freeze panes do not affect named ranges.
        }
        case WorkbookCommand.SetCell _ -> {
          // Value writes do not affect named-range definitions.
        }
        case WorkbookCommand.SetRange _ -> {
          // Matrix writes do not affect named-range definitions.
        }
        case WorkbookCommand.ClearRange _ -> {
          // Clearing cells does not remove or rewrite named-range definitions.
        }
        case WorkbookCommand.SetHyperlink _ -> {
          // Hyperlinks do not affect named ranges.
        }
        case WorkbookCommand.ClearHyperlink _ -> {
          // Hyperlinks do not affect named ranges.
        }
        case WorkbookCommand.SetComment _ -> {
          // Comments do not affect named ranges.
        }
        case WorkbookCommand.ClearComment _ -> {
          // Comments do not affect named ranges.
        }
        case WorkbookCommand.ApplyStyle _ -> {
          // Styling does not affect named ranges.
        }
        case WorkbookCommand.SetNamedRange setNamedRange -> {
          ExcelNamedRangeDefinition definition = setNamedRange.definition();
          ExpectedNamedRange expectedNamedRange =
              new ExpectedNamedRange(
                  definition.scope(), definition.target().refersToFormula(), definition.target());
          expectedNamedRanges.put(new NamedRangeKey(definition.name(), definition.scope()), expectedNamedRange);
        }
        case WorkbookCommand.DeleteNamedRange deleteNamedRange ->
            expectedNamedRanges.remove(
                new NamedRangeKey(deleteNamedRange.name(), deleteNamedRange.scope()));
        case WorkbookCommand.AppendRow _ -> {
          // Appending rows does not affect named-range definitions.
        }
        case WorkbookCommand.AutoSizeColumns _ -> {
          // Auto-sizing does not affect named ranges.
        }
        case WorkbookCommand.EvaluateAllFormulas _ -> {
          // Formula evaluation does not affect named-range definitions.
        }
        case WorkbookCommand.ForceFormulaRecalculationOnOpen _ -> {
          // Force-recalc flags do not affect named ranges.
        }
      }
    }
    return Map.copyOf(expectedNamedRanges);
  }

  private static <T> void renameExpectedSheet(
      Map<String, Map<CellCoordinate, T>> expectedValuesBySheet,
      String sheetName,
      String newSheetName) {
    Map<CellCoordinate, T> expectedValues = expectedValuesBySheet.remove(sheetName);
    if (expectedValues != null) {
      expectedValuesBySheet.put(newSheetName, expectedValues);
    }
  }

  private static void clearExpectedHyperlink(
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadataBySheet,
      String sheetName,
      String address) {
    mutateExpectedMetadata(
        expectedMetadataBySheet,
        sheetName,
        address,
        metadata -> metadata.clearHyperlink());
  }

  private static void clearExpectedComment(
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadataBySheet,
      String sheetName,
      String address) {
    mutateExpectedMetadata(
        expectedMetadataBySheet,
        sheetName,
        address,
        metadata -> metadata.clearComment());
  }

  private static void clearExpectedRange(
      Map<String, Map<CellCoordinate, ExpectedStyle>> expectedStylesBySheet,
      String sheetName,
      String range) {
    Map<CellCoordinate, ExpectedStyle> expectedStyles = expectedStylesBySheet.get(sheetName);
    if (expectedStyles == null) {
      return;
    }
    forEachCell(range, expectedStyles::remove);
  }

  private static void clearExpectedMetadataRange(
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadataBySheet,
      String sheetName,
      String range) {
    Map<CellCoordinate, ExpectedCellMetadata> expectedMetadata = expectedMetadataBySheet.get(sheetName);
    if (expectedMetadata == null) {
      return;
    }
    forEachCell(range, expectedMetadata::remove);
  }

  private static void applyExpectedStyle(
      Map<String, Map<CellCoordinate, ExpectedStyle>> expectedStylesBySheet,
      String sheetName,
      String range,
      ExcelCellStyle style) {
    Map<CellCoordinate, ExpectedStyle> expectedStyles =
        expectedStylesBySheet.computeIfAbsent(sheetName, key -> new HashMap<>());
    forEachCell(
        range,
        coordinate ->
            expectedStyles.merge(
                coordinate,
                ExpectedStyle.from(style),
                ExpectedStyle::merge));
  }

  private static void applyExpectedMetadata(
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadataBySheet,
      String sheetName,
      String address,
      ExpectedCellMetadata metadata) {
    Map<CellCoordinate, ExpectedCellMetadata> expectedMetadata =
        expectedMetadataBySheet.computeIfAbsent(sheetName, key -> new HashMap<>());
    CellReference cellReference = new CellReference(address);
    CellCoordinate coordinate = new CellCoordinate(cellReference.getRow(), cellReference.getCol());
    expectedMetadata.merge(coordinate, metadata, ExpectedCellMetadata::merge);
  }

  private static void mutateExpectedMetadata(
      Map<String, Map<CellCoordinate, ExpectedCellMetadata>> expectedMetadataBySheet,
      String sheetName,
      String address,
      java.util.function.UnaryOperator<ExpectedCellMetadata> mutation) {
    Map<CellCoordinate, ExpectedCellMetadata> expectedMetadata = expectedMetadataBySheet.get(sheetName);
    if (expectedMetadata == null) {
      return;
    }
    CellReference cellReference = new CellReference(address);
    CellCoordinate coordinate = new CellCoordinate(cellReference.getRow(), cellReference.getCol());
    ExpectedCellMetadata existing = expectedMetadata.get(coordinate);
    if (existing == null) {
      return;
    }
    ExpectedCellMetadata updated = mutation.apply(existing);
    if (updated.isEmpty()) {
      expectedMetadata.remove(coordinate);
      return;
    }
    expectedMetadata.put(coordinate, updated);
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

  private static void renameExpectedNamedRanges(
      Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges, String sheetName, String newSheetName) {
    LinkedHashMap<NamedRangeKey, ExpectedNamedRange> renamed = new LinkedHashMap<>();
    expectedNamedRanges.forEach(
        (key, value) -> {
          NamedRangeKey renamedKey = key.renamedSheet(sheetName, newSheetName);
          ExpectedNamedRange renamedValue = value.renamedSheet(sheetName, newSheetName);
          renamed.put(renamedKey, renamedValue);
        });
    expectedNamedRanges.clear();
    expectedNamedRanges.putAll(renamed);
  }

  private static void removeExpectedNamedRangesForSheet(
      Map<NamedRangeKey, ExpectedNamedRange> expectedNamedRanges, String sheetName) {
    expectedNamedRanges.entrySet().removeIf(entry -> entry.getKey().referencesSheet(sheetName)
        || entry.getValue().referencesSheet(sheetName));
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

  private record CellCoordinate(int rowIndex, int columnIndex) {
    private String a1Address() {
      return new CellReference(rowIndex, columnIndex).formatAsString();
    }
  }

  private record ExpectedCellMetadata(ExcelHyperlink hyperlink, ExcelComment comment) {
    private static ExpectedCellMetadata withHyperlink(ExcelHyperlink hyperlink) {
      return new ExpectedCellMetadata(Objects.requireNonNull(hyperlink, "hyperlink must not be null"), null);
    }

    private static ExpectedCellMetadata withComment(ExcelComment comment) {
      return new ExpectedCellMetadata(null, Objects.requireNonNull(comment, "comment must not be null"));
    }

    private ExpectedCellMetadata merge(ExpectedCellMetadata overlay) {
      return new ExpectedCellMetadata(
          overlay.hyperlink() == null ? hyperlink : overlay.hyperlink(),
          overlay.comment() == null ? comment : overlay.comment());
    }

    private ExpectedCellMetadata clearHyperlink() {
      return new ExpectedCellMetadata(null, comment);
    }

    private ExpectedCellMetadata clearComment() {
      return new ExpectedCellMetadata(hyperlink, null);
    }

    private boolean isEmpty() {
      return hyperlink == null && comment == null;
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

    private NamedRangeKey renamedSheet(String sheetName, String newSheetName) {
      return switch (scope) {
        case ExcelNamedRangeScope.WorkbookScope _ -> this;
        case ExcelNamedRangeScope.SheetScope sheetScope ->
            sheetScope.sheetName().equals(sheetName)
                ? new NamedRangeKey(name, new ExcelNamedRangeScope.SheetScope(newSheetName))
                : this;
      };
    }

    private boolean referencesSheet(String sheetName) {
      return switch (scope) {
        case ExcelNamedRangeScope.WorkbookScope _ -> false;
        case ExcelNamedRangeScope.SheetScope sheetScope -> sheetScope.sheetName().equals(sheetName);
      };
    }
  }

  private record ExpectedNamedRange(
      ExcelNamedRangeScope scope, String refersToFormula, ExcelNamedRangeTarget target) {
    private ExpectedNamedRange renamedSheet(String sheetName, String newSheetName) {
      ExcelNamedRangeScope renamedScope =
          switch (scope) {
            case ExcelNamedRangeScope.WorkbookScope _ -> scope;
            case ExcelNamedRangeScope.SheetScope sheetScope ->
                sheetScope.sheetName().equals(sheetName)
                    ? new ExcelNamedRangeScope.SheetScope(newSheetName)
                    : scope;
          };
      ExcelNamedRangeTarget renamedTarget =
          target != null && target.sheetName().equals(sheetName)
              ? new ExcelNamedRangeTarget(newSheetName, target.range())
              : target;
      String renamedFormula =
          renamedTarget == null
              ? refersToFormula
              : renamedTarget.refersToFormula();
      return new ExpectedNamedRange(renamedScope, renamedFormula, renamedTarget);
    }

    private boolean referencesSheet(String sheetName) {
      boolean scopeReferencesSheet =
          switch (scope) {
            case ExcelNamedRangeScope.WorkbookScope _ -> false;
            case ExcelNamedRangeScope.SheetScope sheetScope -> sheetScope.sheetName().equals(sheetName);
          };
      return scopeReferencesSheet || (target != null && target.sheetName().equals(sheetName));
    }
  }

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
    private static ExpectedStyle from(ExcelCellStyle style) {
      return new ExpectedStyle(
          style.numberFormat(),
          style.bold(),
          style.italic(),
          style.wrapText(),
          style.horizontalAlignment(),
          style.verticalAlignment(),
          style.fontName(),
          style.fontHeight() == null ? null : style.fontHeight().twips(),
          style.fontColor(),
          style.underline(),
          style.strikeout(),
          style.fillColor(),
          borderStyle(style.border(), BorderEdge.TOP),
          borderStyle(style.border(), BorderEdge.RIGHT),
          borderStyle(style.border(), BorderEdge.BOTTOM),
          borderStyle(style.border(), BorderEdge.LEFT));
    }

    private ExpectedStyle merge(ExpectedStyle overlay) {
      return new ExpectedStyle(
          overlay.numberFormat() == null ? numberFormat : overlay.numberFormat(),
          overlay.bold() == null ? bold : overlay.bold(),
          overlay.italic() == null ? italic : overlay.italic(),
          overlay.wrapText() == null ? wrapText : overlay.wrapText(),
          overlay.horizontalAlignment() == null ? horizontalAlignment : overlay.horizontalAlignment(),
          overlay.verticalAlignment() == null ? verticalAlignment : overlay.verticalAlignment(),
          overlay.fontName() == null ? fontName : overlay.fontName(),
          overlay.fontHeightTwips() == null ? fontHeightTwips : overlay.fontHeightTwips(),
          overlay.fontColor() == null ? fontColor : overlay.fontColor(),
          overlay.underline() == null ? underline : overlay.underline(),
          overlay.strikeout() == null ? strikeout : overlay.strikeout(),
          overlay.fillColor() == null ? fillColor : overlay.fillColor(),
          overlay.borderTop() == null ? borderTop : overlay.borderTop(),
          overlay.borderRight() == null ? borderRight : overlay.borderRight(),
          overlay.borderBottom() == null ? borderBottom : overlay.borderBottom(),
          overlay.borderLeft() == null ? borderLeft : overlay.borderLeft());
    }

    private static ExcelBorderStyle borderStyle(ExcelBorder border, BorderEdge edge) {
      if (border == null) {
        return null;
      }
      return switch (edge) {
        case TOP -> borderStyle(border.all(), border.top());
        case RIGHT -> borderStyle(border.all(), border.right());
        case BOTTOM -> borderStyle(border.all(), border.bottom());
        case LEFT -> borderStyle(border.all(), border.left());
      };
    }

    private static ExcelBorderStyle borderStyle(
        ExcelBorderSide defaultSide, ExcelBorderSide explicitSide) {
      if (explicitSide != null) {
        return explicitSide.style();
      }
      return defaultSide == null ? null : defaultSide.style();
    }
  }

  private enum BorderEdge {
    TOP,
    RIGHT,
    BOTTOM,
    LEFT
  }
}
