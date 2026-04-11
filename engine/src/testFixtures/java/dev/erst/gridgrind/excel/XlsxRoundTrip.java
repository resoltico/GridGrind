package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientStop;

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

  /** Returns the active sheet name exactly as stored in the saved workbook. */
  public static String activeSheetName(Path workbookPath) throws IOException {
    return readWorkbook(
        workbookPath,
        workbook -> {
          int activeSheetIndex = workbook.getActiveSheetIndex();
          if (activeSheetIndex < 0 || activeSheetIndex >= workbook.getNumberOfSheets()) {
            throw new IllegalStateException("active sheet index must reference one saved sheet");
          }
          return workbook.getSheetName(activeSheetIndex);
        });
  }

  /** Returns the selected sheet names in workbook order exactly as stored in the saved workbook. */
  public static List<String> selectedSheetNames(Path workbookPath) throws IOException {
    return readWorkbook(
        workbookPath,
        workbook -> {
          List<String> selectedSheetNames = new ArrayList<>();
          for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            if (workbook.getSheetAt(sheetIndex).isSelected()) {
              selectedSheetNames.add(workbook.getSheetName(sheetIndex));
            }
          }
          return List.copyOf(selectedSheetNames);
        });
  }

  /** Returns the saved visibility state for one sheet. */
  public static ExcelSheetVisibility sheetVisibility(Path workbookPath, String sheetName)
      throws IOException {
    return readWorkbook(
        workbookPath,
        workbook -> {
          int sheetIndex = workbook.getSheetIndex(sheetName);
          if (sheetIndex < 0) {
            throw new SheetNotFoundException(sheetName);
          }
          return ExcelSheetVisibility.fromPoi(workbook.getSheetVisibility(sheetIndex));
        });
  }

  /** Returns the saved sheet-protection state for one sheet. */
  public static WorkbookReadResult.SheetProtection sheetProtection(
      Path workbookPath, String sheetName) throws IOException {
    return readSheet(
        workbookPath,
        sheetName,
        sheet -> {
          if (!sheet.getProtect()) {
            return new WorkbookReadResult.SheetProtection.Unprotected();
          }
          return new WorkbookReadResult.SheetProtection.Protected(
              new ExcelSheetProtectionSettings(
                  sheet.isAutoFilterLocked(),
                  sheet.isDeleteColumnsLocked(),
                  sheet.isDeleteRowsLocked(),
                  sheet.isFormatCellsLocked(),
                  sheet.isFormatColumnsLocked(),
                  sheet.isFormatRowsLocked(),
                  sheet.isInsertColumnsLocked(),
                  sheet.isInsertHyperlinksLocked(),
                  sheet.isInsertRowsLocked(),
                  sheet.isObjectsLocked(),
                  sheet.isPivotTablesLocked(),
                  sheet.isScenariosLocked(),
                  sheet.isSelectLockedCellsLocked(),
                  sheet.isSelectUnlockedCellsLocked(),
                  sheet.isSortLocked()));
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

  /** Returns the explicit pane state stored on the named sheet. */
  public static ExcelSheetPane pane(Path workbookPath, String sheetName) throws IOException {
    return readSheet(workbookPath, sheetName, ExcelSheetViewSupport::pane);
  }

  /** Returns the effective zoom percentage stored on the named sheet. */
  public static int zoomPercent(Path workbookPath, String sheetName) throws IOException {
    return readSheet(workbookPath, sheetName, ExcelSheetViewSupport::zoomPercent);
  }

  /** Returns the supported sheet-layout snapshot stored on the named sheet. */
  public static WorkbookReadResult.SheetLayout sheetLayout(Path workbookPath, String sheetName)
      throws IOException {
    requireWorkbookPath(workbookPath);
    requireNonBlank(sheetName, "sheetName");

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      return ((WorkbookReadResult.SheetLayoutResult)
              new WorkbookReadExecutor()
                  .apply(workbook, new WorkbookReadCommand.GetSheetLayout("layout", sheetName))
                  .getFirst())
          .layout();
    }
  }

  /** Returns the supported print-layout state stored on the named sheet. */
  public static ExcelPrintLayout printLayout(Path workbookPath, String sheetName)
      throws IOException {
    ExcelPrintLayoutController controller = new ExcelPrintLayoutController();
    return readSheet(workbookPath, sheetName, controller::printLayout);
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
                  colorSnapshot(font.getXSSFColor()),
                  font.getUnderline() != org.apache.poi.ss.usermodel.Font.U_NONE,
                  font.getStrikeout()),
              fillSnapshot((XSSFWorkbook) sheet.getWorkbook(), style),
              new ExcelBorderSnapshot(
                  borderSideSnapshot(style.getBorderTop(), style.getTopBorderXSSFColor()),
                  borderSideSnapshot(style.getBorderRight(), style.getRightBorderXSSFColor()),
                  borderSideSnapshot(style.getBorderBottom(), style.getBottomBorderXSSFColor()),
                  borderSideSnapshot(style.getBorderLeft(), style.getLeftBorderXSSFColor())),
              new ExcelCellProtectionSnapshot(style.getLocked(), style.getHidden()));
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
          return ExcelCellMetadataSnapshot.of(hyperlink(cell), commentSnapshot(cell));
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

  private static ExcelCellFillSnapshot fillSnapshot(XSSFWorkbook workbook, XSSFCellStyle style) {
    XSSFCellFill fill = fill(workbook, style);
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
        colorSnapshot(style.getFillForegroundColorColor()),
        pattern == ExcelFillPattern.SOLID
            ? null
            : colorSnapshot(style.getFillBackgroundColorColor()));
  }

  private static XSSFCellFill fill(XSSFWorkbook workbook, XSSFCellStyle style) {
    return workbook.getStylesSource().getFillAt((int) style.getCoreXf().getFillId());
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

  private static ExcelCommentSnapshot commentSnapshot(XSSFCell cell) {
    var comment = cell.getCellComment();
    if (comment == null || comment.getString() == null) {
      return null;
    }
    String text = comment.getString().getString();
    String author = comment.getAuthor();
    if (text == null || text.isBlank() || author == null || author.isBlank()) {
      return null;
    }
    return ExcelSheet.commentSnapshot(cell);
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

  private static ExcelColorSnapshot colorSnapshot(XSSFColor color) {
    if (color == null) {
      return null;
    }
    byte[] rgb = color.getRGB();
    String rgbHex =
        rgb == null || rgb.length != 3
            ? null
            : "#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    Integer theme = color.isThemed() ? color.getTheme() : null;
    Integer indexed = color.isIndexed() ? Short.toUnsignedInt(color.getIndexed()) : null;
    Double tint = color.hasTint() ? color.getTint() : null;
    if (rgbHex == null && theme == null && indexed == null) {
      return null;
    }
    return new ExcelColorSnapshot(rgbHex, theme, indexed, tint);
  }

  private static ExcelGradientFillSnapshot gradientFillSnapshot(
      XSSFWorkbook workbook, CTGradientFill fill) {
    List<ExcelGradientStopSnapshot> stops =
        java.util.Arrays.stream(fill.getStopArray())
            .map(stop -> gradientStopSnapshot(workbook, stop))
            .toList();
    return new ExcelGradientFillSnapshot(
        fill.isSetType() ? fill.getType().toString() : "LINEAR",
        fill.isSetDegree() ? fill.getDegree() : null,
        fill.isSetLeft() ? fill.getLeft() : null,
        fill.isSetRight() ? fill.getRight() : null,
        fill.isSetTop() ? fill.getTop() : null,
        fill.isSetBottom() ? fill.getBottom() : null,
        stops);
  }

  private static ExcelGradientStopSnapshot gradientStopSnapshot(
      XSSFWorkbook workbook, CTGradientStop stop) {
    return new ExcelGradientStopSnapshot(
        stop.getPosition(), ExcelColorSnapshotSupport.snapshot(workbook, stop.getColor()));
  }

  private static ExcelBorderSideSnapshot borderSideSnapshot(
      BorderStyle borderStyle, XSSFColor borderColor) {
    ExcelBorderStyle style = fromPoi(borderStyle);
    return new ExcelBorderSideSnapshot(
        style, style == ExcelBorderStyle.NONE ? null : colorSnapshot(borderColor));
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

  private static ExcelFillPattern fromPoi(org.apache.poi.ss.usermodel.FillPatternType pattern) {
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
}
