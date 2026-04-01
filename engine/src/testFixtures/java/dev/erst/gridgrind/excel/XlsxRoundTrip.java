package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Creates temporary workbook paths and reopens saved `.xlsx` files for structural inspection. */
public final class XlsxRoundTrip {
  private XlsxRoundTrip() {}

  /** Returns a fresh temporary path for a workbook that does not yet exist on disk. */
  public static Path newWorkbookPath(String prefix) throws IOException {
    requireNonBlank(prefix, "prefix");
    Path directory = Files.createTempDirectory(prefix);
    return directory.resolve("workbook.xlsx").toAbsolutePath();
  }

  /** Returns the sheet order exactly as stored in the saved workbook. */
  public static List<String> sheetOrder(Path workbookPath) throws IOException {
    return readWorkbook(
        workbookPath,
        workbook -> {
          List<String> sheetOrder = new ArrayList<>(workbook.getNumberOfSheets());
          for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            sheetOrder.add(workbook.getSheetName(sheetIndex));
          }
          return List.copyOf(sheetOrder);
        });
  }

  /** Returns all merged regions on the named sheet using A1-style range strings. */
  public static List<String> mergedRegions(Path workbookPath, String sheetName) throws IOException {
    return readSheet(
        workbookPath,
        sheetName,
        sheet -> {
          List<String> mergedRegions = new ArrayList<>(sheet.getNumMergedRegions());
          for (int regionIndex = 0; regionIndex < sheet.getNumMergedRegions(); regionIndex++) {
            mergedRegions.add(sheet.getMergedRegion(regionIndex).formatAsString());
          }
          return List.copyOf(mergedRegions);
        });
  }

  /** Returns the requested column width in Apache POI's raw width units. */
  public static int columnWidth(Path workbookPath, String sheetName, int columnIndex)
      throws IOException {
    requireNonNegative(columnIndex, "columnIndex");
    return readSheet(workbookPath, sheetName, sheet -> sheet.getColumnWidth(columnIndex));
  }

  /** Returns the requested row height in twips, falling back to the sheet default when absent. */
  public static short rowHeightTwips(Path workbookPath, String sheetName, int rowIndex)
      throws IOException {
    requireNonNegative(rowIndex, "rowIndex");
    return readSheet(
        workbookPath,
        sheetName,
        sheet -> {
          Row row = sheet.getRow(rowIndex);
          return row == null ? sheet.getDefaultRowHeight() : row.getHeight();
        });
  }

  /** Returns the freeze-pane state stored on the named sheet. */
  public static FreezePaneState freezePaneState(Path workbookPath, String sheetName)
      throws IOException {
    return readSheet(
        workbookPath,
        sheetName,
        sheet -> {
          PaneInformation paneInformation = sheet.getPaneInformation();
          if (paneInformation == null || !paneInformation.isFreezePane()) {
            return new FreezePaneState.None();
          }
          return new FreezePaneState.Frozen(
              paneInformation.getVerticalSplitPosition(),
              paneInformation.getHorizontalSplitPosition(),
              paneInformation.getVerticalSplitLeftColumn(),
              paneInformation.getHorizontalSplitTopRow());
        });
  }

  /** Returns the effective style captured at one A1-style cell address in the saved workbook. */
  public static ExcelCellStyleSnapshot cellStyle(
      Path workbookPath, String sheetName, String address) throws IOException {
    requireNonBlank(address, "address");
    CellReference cellReference = new CellReference(address);
    return readSheet(
        workbookPath,
        sheetName,
        sheet -> {
          Row row = sheet.getRow(cellReference.getRow());
          if (row == null) {
            throw new CellNotFoundException(address);
          }
          XSSFCell cell = (XSSFCell) row.getCell(cellReference.getCol());
          if (cell == null) {
            throw new CellNotFoundException(address);
          }
          XSSFCellStyle style = cell.getCellStyle();
          XSSFFont font = style.getFont();
          return new ExcelCellStyleSnapshot(
              WorkbookStyleRegistry.resolveNumberFormat(style.getDataFormatString()),
              font.getBold(),
              font.getItalic(),
              style.getWrapText(),
              fromPoi(style.getAlignment()),
              fromPoi(style.getVerticalAlignment()),
              font.getFontName(),
              new ExcelFontHeight(font.getFontHeight()),
              toRgbHex(font.getXSSFColor()),
              font.getUnderline() != org.apache.poi.ss.usermodel.Font.U_NONE,
              font.getStrikeout(),
              fillColor(style),
              fromPoi(style.getBorderTop()),
              fromPoi(style.getBorderRight()),
              fromPoi(style.getBorderBottom()),
              fromPoi(style.getBorderLeft()));
        });
  }

  /** Returns the optional hyperlink and comment facts stored at one saved cell address. */
  public static ExcelCellMetadataSnapshot cellMetadata(
      Path workbookPath, String sheetName, String address) throws IOException {
    requireNonBlank(address, "address");
    CellReference cellReference = new CellReference(address);
    return readSheet(
        workbookPath,
        sheetName,
        sheet -> {
          Row row = sheet.getRow(cellReference.getRow());
          if (row == null) {
            throw new CellNotFoundException(address);
          }
          XSSFCell cell = (XSSFCell) row.getCell(cellReference.getCol());
          if (cell == null) {
            throw new CellNotFoundException(address);
          }
          return ExcelCellMetadataSnapshot.of(hyperlink(cell), comment(cell));
        });
  }

  /** Returns every analyzable named range stored in the saved workbook. */
  public static List<ExcelNamedRangeSnapshot> namedRanges(Path workbookPath) throws IOException {
    return readWorkbook(
        workbookPath,
        workbook -> {
          List<ExcelNamedRangeSnapshot> namedRanges = new ArrayList<>();
          for (var name : workbook.getAllNames()) {
            if (!shouldExpose(name)) {
              continue;
            }
            ExcelNamedRangeScope scope = toScope(workbook, name.getSheetIndex());
            String refersToFormula = Objects.requireNonNullElse(name.getRefersToFormula(), "");
            var target = ExcelNamedRangeTargets.resolveTarget(refersToFormula, scope);
            if (target.isEmpty()) {
              namedRanges.add(
                  new ExcelNamedRangeSnapshot.FormulaSnapshot(
                      name.getNameName(), scope, refersToFormula));
            } else {
              namedRanges.add(
                  new ExcelNamedRangeSnapshot.RangeSnapshot(
                      name.getNameName(), scope, refersToFormula, target.orElseThrow()));
            }
          }
          return List.copyOf(namedRanges);
        });
  }

  /** Returns normalized data-validation structures stored on one saved sheet. */
  public static List<ExcelDataValidationSnapshot> dataValidations(
      Path workbookPath, String sheetName) throws IOException {
    requireWorkbookPath(workbookPath);
    requireNonBlank(sheetName, "sheetName");

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      return workbook.sheet(sheetName).dataValidations(new ExcelRangeSelection.All());
    }
  }

  private static <T> T readSheet(Path workbookPath, String sheetName, SheetReader<T> sheetReader)
      throws IOException {
    requireNonBlank(sheetName, "sheetName");
    return readWorkbook(
        workbookPath,
        workbook -> {
          XSSFSheet sheet = workbook.getSheet(sheetName);
          if (sheet == null) {
            throw new SheetNotFoundException(sheetName);
          }
          return sheetReader.read(sheet);
        });
  }

  private static <T> T readWorkbook(Path workbookPath, WorkbookReader<T> workbookReader)
      throws IOException {
    Path absolutePath = requireWorkbookPath(workbookPath);

    try (InputStream inputStream = Files.newInputStream(absolutePath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
      return workbookReader.read(workbook);
    }
  }

  private static Path requireWorkbookPath(Path workbookPath) {
    Objects.requireNonNull(workbookPath, "workbookPath must not be null");

    Path absolutePath = workbookPath.toAbsolutePath();
    if (!Files.exists(absolutePath)) {
      throw new WorkbookNotFoundException(absolutePath);
    }
    return absolutePath;
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  /** Reads a saved workbook and returns a derived value. */
  @FunctionalInterface
  private interface WorkbookReader<T> {
    /** Reads the workbook and returns the derived value. */
    T read(XSSFWorkbook workbook) throws IOException;
  }

  /** Reads a saved sheet and returns a derived value. */
  @FunctionalInterface
  private interface SheetReader<T> {
    /** Reads the sheet and returns the derived value. */
    T read(XSSFSheet sheet);
  }

  private static String fillColor(XSSFCellStyle style) {
    if (style.getFillPattern() != FillPatternType.SOLID_FOREGROUND) {
      return null;
    }
    return toRgbHex(style.getFillForegroundColorColor());
  }

  private static ExcelHyperlink hyperlink(XSSFCell cell) {
    XSSFHyperlink hyperlink = cell.getHyperlink();
    if (hyperlink == null || hyperlink.getType() == null) {
      return null;
    }
    String address = hyperlink.getAddress();
    if (address == null || address.isBlank()) {
      return null;
    }
    try {
      return switch (hyperlink.getType()) {
        case URL -> new ExcelHyperlink.Url(address);
        case EMAIL -> new ExcelHyperlink.Email(address);
        case FILE -> new ExcelHyperlink.File(address);
        case DOCUMENT -> new ExcelHyperlink.Document(address);
        case NONE -> null;
      };
    } catch (IllegalArgumentException exception) {
      return null;
    }
  }

  private static ExcelComment comment(XSSFCell cell) {
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

  private static boolean shouldExpose(org.apache.poi.ss.usermodel.Name name) {
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

  private static ExcelHorizontalAlignment fromPoi(HorizontalAlignment alignment) {
    return ExcelHorizontalAlignment.valueOf(alignment.name());
  }

  private static ExcelVerticalAlignment fromPoi(VerticalAlignment alignment) {
    return ExcelVerticalAlignment.valueOf(alignment.name());
  }

  private static ExcelBorderStyle fromPoi(BorderStyle borderStyle) {
    return ExcelBorderStyle.valueOf(borderStyle.name());
  }

  /** Represents the freeze-pane state stored on a sheet. */
  public sealed interface FreezePaneState permits FreezePaneState.None, FreezePaneState.Frozen {
    /** Represents a sheet without any frozen panes. */
    record None() implements FreezePaneState {}

    /** Represents a sheet with a frozen pane defined by split and visible-origin coordinates. */
    record Frozen(int splitColumn, int splitRow, int leftmostColumn, int topRow)
        implements FreezePaneState {
      /** Validates the stored freeze-pane coordinates. */
      public Frozen {
        requireNonNegative(splitColumn, "splitColumn");
        requireNonNegative(splitRow, "splitRow");
        requireNonNegative(leftmostColumn, "leftmostColumn");
        requireNonNegative(topRow, "topRow");
      }
    }
  }
}
