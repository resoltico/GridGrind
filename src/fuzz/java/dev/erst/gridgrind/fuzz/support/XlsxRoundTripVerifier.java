package dev.erst.gridgrind.fuzz.support;

import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientStopReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelGradientFillGeometry;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.List;
import java.util.Optional;
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
      XssfWorkbook workbook, Path workbookPath, List<ResolvedStep> commands) throws IOException {
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

  static CellStyleReport defaultStyleSnapshot() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Default");
      return workbook.sheet("Default").cells().snapshotCell("A1").style().orElseThrow();
    }
  }

  static CellStyleReport styleSnapshot(XSSFWorkbook workbook, XSSFCellStyle style) {
    XSSFFont font = style.getFont();
    return new CellStyleReport(
        resolveNumberFormat(style.getDataFormatString()),
        new CellAlignmentReport(
            style.getWrapText(),
            fromPoi(style.getAlignment()),
            fromPoi(style.getVerticalAlignment()),
            style.getRotation(),
            style.getIndention()),
        new CellFontReport(
            font.getBold(),
            font.getItalic(),
            font.getFontName(),
            FontHeightReport.fromTwips(font.getFontHeight()),
            toColorSnapshot(font.getXSSFColor()),
            font.getUnderline() != org.apache.poi.ss.usermodel.Font.U_NONE,
            font.getStrikeout()),
        fillSnapshot(workbook, style),
        new CellBorderReport(
            borderSide(style.getBorderTop(), style.getTopBorderXSSFColor()),
            borderSide(style.getBorderRight(), style.getRightBorderXSSFColor()),
            borderSide(style.getBorderBottom(), style.getBottomBorderXSSFColor()),
            borderSide(style.getBorderLeft(), style.getLeftBorderXSSFColor())),
        new CellProtectionReport(style.getLocked(), style.getHidden()));
  }

  private static CellBorderSideReport borderSide(BorderStyle border, XSSFColor color) {
    ExcelBorderStyle style = fromPoi(border);
    if (style == ExcelBorderStyle.NONE) {
      return new CellBorderSideReport.None();
    }
    return toColorSnapshot(color)
        .<CellBorderSideReport>map(value -> new CellBorderSideReport.Colored(style, value))
        .orElseGet(() -> new CellBorderSideReport.DefaultColor(style));
  }

  private static CellFillReport fillSnapshot(XSSFWorkbook workbook, XSSFCellStyle style) {
    XSSFCellFill fill = workbook.getStylesSource().getFillAt((int) style.getCoreXf().getFillId());
    if (fill.getCTFill().isSetGradientFill()) {
      return CellFillReport.gradient(
          gradientFillSnapshot(workbook, fill.getCTFill().getGradientFill()));
    }
    ExcelFillPattern pattern = fromPoi(style.getFillPattern());
    if (pattern == ExcelFillPattern.NONE) {
      return CellFillReport.pattern(pattern);
    }
    Optional<CellColorReport> foreground = toColorSnapshot(style.getFillForegroundColorColor());
    CellColorReport background =
        pattern == ExcelFillPattern.SOLID
            ? null
            : toColorSnapshot(style.getFillBackgroundColorColor()).orElse(null);
    if (foreground.isPresent() && background != null) {
      return CellFillReport.patternColors(pattern, foreground.orElseThrow(), background);
    }
    if (foreground.isPresent()) {
      return CellFillReport.patternForeground(pattern, foreground.orElseThrow());
    }
    if (background != null) {
      return CellFillReport.patternBackground(pattern, background);
    }
    return CellFillReport.pattern(pattern);
  }

  private static CellGradientFillReport gradientFillSnapshot(
      XSSFWorkbook workbook, CTGradientFill fill) {
    Double left = fill.isSetLeft() ? fill.getLeft() : null;
    Double right = fill.isSetRight() ? fill.getRight() : null;
    Double top = fill.isSetTop() ? fill.getTop() : null;
    Double bottom = fill.isSetBottom() ? fill.getBottom() : null;
    List<CellGradientStopReport> stops =
        java.util.Arrays.stream(fill.getStopArray())
            .map(stop -> gradientStopSnapshot(workbook, stop))
            .toList();
    String type =
        ExcelGradientFillGeometry.effectiveType(
            fill.isSetType() ? fill.getType().toString() : null, left, right, top, bottom);
    return "PATH".equals(type)
        ? CellGradientFillReport.path(
            Optional.ofNullable(left),
            Optional.ofNullable(right),
            Optional.ofNullable(top),
            Optional.ofNullable(bottom),
            stops)
        : CellGradientFillReport.linear(
            Optional.ofNullable(fill.isSetDegree() ? fill.getDegree() : null), stops);
  }

  private static CellGradientStopReport gradientStopSnapshot(
      XSSFWorkbook workbook, CTGradientStop stop) {
    XSSFColor color =
        XSSFColor.from(stop.getColor(), workbook.getStylesSource().getIndexedColors());
    ThemesTable themes = workbook.getStylesSource().getTheme();
    if (themes != null) {
      themes.inheritFromThemeAsRequired(color);
    }
    CellColorReport snapshot =
        java.util.Objects.requireNonNull(
            toColorSnapshot(color).orElse(null),
            "gradient stop color must resolve to one snapshot");
    return new CellGradientStopReport(stop.getPosition(), snapshot);
  }

  private static void requireColorShape(XSSFColor color, String label) {
    if (color == null || color.getRGB() == null) {
      return;
    }
    if (color.getRGB().length != 3) {
      throw new IllegalStateException(label + " must resolve to a 3-byte RGB value");
    }
  }

  private static Optional<CellColorReport> toColorSnapshot(XSSFColor color) {
    if (color == null) {
      return Optional.empty();
    }
    byte[] rgb = color.getRGB();
    String rgbHex = null;
    if (rgb != null) {
      if (rgb.length != 3) {
        return Optional.empty();
      }
      rgbHex = "#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }
    Integer theme = color.isThemed() ? color.getTheme() : null;
    Integer indexed = color.isIndexed() ? Short.toUnsignedInt(color.getIndexed()) : null;
    Optional<Double> tint = color.hasTint() ? Optional.of(color.getTint()) : Optional.empty();
    if (theme != null || indexed != null) {
      rgbHex = null;
    }
    if (rgbHex == null && theme == null && indexed == null) {
      return Optional.empty();
    }
    if (rgbHex != null) {
      return Optional.of(new CellColorReport.Rgb(new RgbHex(rgbHex), tint));
    }
    if (theme != null) {
      return Optional.of(new CellColorReport.Theme(theme, tint));
    }
    return Optional.of(
        new CellColorReport.Indexed(
            java.util.Objects.requireNonNull(indexed, "indexed color must be present"), tint));
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

  static Optional<HyperlinkTarget> hyperlink(Cell cell) {
    XSSFHyperlink hyperlink = (XSSFHyperlink) cell.getHyperlink();
    if (hyperlink == null || hyperlink.getType() == null) {
      return Optional.empty();
    }
    String target = hyperlink.getAddress();
    if (target == null || target.isBlank()) {
      return Optional.empty();
    }
    try {
      return switch (hyperlink.getType()) {
        case URL -> Optional.of(new HyperlinkTarget.Url(target));
        case EMAIL -> Optional.of(new HyperlinkTarget.Email(target));
        case FILE -> Optional.of(new HyperlinkTarget.File(target));
        case DOCUMENT -> Optional.of(new HyperlinkTarget.Document(target));
        case NONE -> Optional.empty();
      };
    } catch (IllegalArgumentException exception) {
      return Optional.empty();
    }
  }

  static Optional<ResolvedComment> comment(Cell cell) {
    var comment = cell.getCellComment();
    if (comment == null || comment.getString() == null) {
      return Optional.empty();
    }
    String text = comment.getString().getString();
    String author = comment.getAuthor();
    if (text == null || text.isBlank() || author == null || author.isBlank()) {
      return Optional.empty();
    }
    return Optional.of(new ResolvedComment(text, author, comment.isVisible()));
  }
}
