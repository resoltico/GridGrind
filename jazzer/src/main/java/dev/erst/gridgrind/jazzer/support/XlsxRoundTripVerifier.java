package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSnapshot;
import dev.erst.gridgrind.excel.ExcelCellAlignmentSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellProtectionSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelColorSnapshot;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelGradientFillSnapshot;
import dev.erst.gridgrind.excel.ExcelGradientStopSnapshot;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelGradientFillGeometry;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.List;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
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
    XlsxRoundTripExpectedStateSupport.ExpectedWorkbookState expectedWorkbookState =
        XlsxRoundTripExpectedStateSupport.expectedWorkbookState(workbook, commands);
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
      XlsxRoundTripExpectationVerification.requireExpectedStyles(
          expectedWorkbookState.expectedStyles(), reopenedWorkbook);
      XlsxRoundTripExpectationVerification.requireExpectedMetadata(
          expectedWorkbookState.expectedMetadata(), reopenedWorkbook);
      XlsxRoundTripExpectationVerification.requireExpectedNamedRanges(
          expectedWorkbookState.expectedNamedRanges(), reopenedWorkbook);
    }
    XlsxRoundTripExpectationVerification.requireExpectedWorkbookState(
        expectedWorkbookState, workbookPath);
    XlsxRoundTripExpectationVerification.requireExpectedSheetLayouts(
        expectedWorkbookState.expectedSheetLayouts(), workbookPath);
    XlsxRoundTripExpectationVerification.requireExpectedDataValidations(
        expectedWorkbookState.expectedDataValidations(), workbookPath);
    XlsxRoundTripExpectationVerification.requireExpectedConditionalFormatting(
        expectedWorkbookState.expectedConditionalFormatting(), workbookPath);
    XlsxRoundTripExpectationVerification.requireExpectedAutofilters(
        expectedWorkbookState.expectedAutofilters(), workbookPath);
    XlsxRoundTripExpectationVerification.requireExpectedTables(
        expectedWorkbookState.expectedTables(), workbookPath);
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

  static String resolveNumberFormat(String numberFormat) {
    return numberFormat == null || numberFormat.isBlank() ? "General" : numberFormat;
  }

  static ExcelCellStyleSnapshot defaultStyleSnapshot() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Default");
      return workbook.sheet("Default").snapshotCell("A1").style();
    }
  }

  static ExcelCellStyleSnapshot styleSnapshot(XSSFWorkbook workbook, XSSFCellStyle style) {
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
      return ExcelCellFillSnapshot.gradient(
          gradientFillSnapshot(workbook, fill.getCTFill().getGradientFill()));
    }
    ExcelFillPattern pattern = fromPoi(style.getFillPattern());
    if (pattern == ExcelFillPattern.NONE) {
      return ExcelCellFillSnapshot.pattern(pattern);
    }
    ExcelColorSnapshot foreground = toColorSnapshot(style.getFillForegroundColorColor());
    ExcelColorSnapshot background =
        pattern == ExcelFillPattern.SOLID
            ? null
            : toColorSnapshot(style.getFillBackgroundColorColor());
    if (foreground != null && background != null) {
      return ExcelCellFillSnapshot.patternColors(pattern, foreground, background);
    }
    if (foreground != null) {
      return ExcelCellFillSnapshot.patternForeground(pattern, foreground);
    }
    if (background != null) {
      return ExcelCellFillSnapshot.patternBackground(pattern, background);
    }
    return ExcelCellFillSnapshot.pattern(pattern);
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
    String type =
        ExcelGradientFillGeometry.effectiveType(
            fill.isSetType() ? fill.getType().toString() : null, left, right, top, bottom);
    return "PATH".equals(type)
        ? ExcelGradientFillSnapshot.path(left, right, top, bottom, stops)
        : ExcelGradientFillSnapshot.linear(fill.isSetDegree() ? fill.getDegree() : null, stops);
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
    if (rgbHex != null) {
      return ExcelColorSnapshot.rgb(rgbHex, tint);
    }
    if (theme != null) {
      return ExcelColorSnapshot.theme(theme, tint);
    }
    return ExcelColorSnapshot.indexed(indexed, tint);
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

  static ExcelHyperlink hyperlink(Cell cell) {
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

  static ExcelComment comment(Cell cell) {
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
}
