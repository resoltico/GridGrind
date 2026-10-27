package dev.erst.gridgrind.engine.styles;

import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillSpec;
import dev.erst.gridgrind.core.model.style.CellGradientStopReport;
import dev.erst.gridgrind.core.model.style.CellGradientStopSpec;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import dev.erst.gridgrind.core.model.style.ResolvedRichTextRun;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelGradientFillGeometry;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.engine.cells.CellTextLimits;
import dev.erst.gridgrind.engine.poi.EnumMappings;
import dev.erst.gridgrind.engine.poi.FillRegistryAccess;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.function.Supplier;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTGradientStop;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRPrElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues;

/** Owns workbook style/font/fill reuse, patch application, style readback and rich-text runs. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
public final class StyleAdapter {
  private static final CellStylePatch LOCAL_DATE_STYLE_PATCH =
      new CellStylePatch(
          Optional.of("yyyy-mm-dd"),
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          Optional.empty());
  private static final CellStylePatch LOCAL_DATE_TIME_STYLE_PATCH =
      new CellStylePatch(
          Optional.of("yyyy-mm-dd hh:mm:ss"),
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          Optional.empty());

  private final XSSFWorkbook workbook;
  private final DataFormat dataFormat;
  private final FillRegistryAccess fillRegistryAccess;
  private final Map<MergedCellStyleKey, XSSFCellStyle> cellStyles;
  private final Map<MergedFontKey, XSSFFont> fonts;
  private final Map<String, Integer> gradientFillIds;

  public StyleAdapter(XSSFWorkbook workbook) {
    this(workbook, FillRegistryAccess.poiApi());
  }

  public StyleAdapter(XSSFWorkbook workbook, FillRegistryAccess fillRegistryAccess) {
    this.workbook = workbook;
    this.fillRegistryAccess = fillRegistryAccess;
    this.dataFormat = workbook.createDataFormat();
    this.cellStyles = new HashMap<>();
    this.fonts = new HashMap<>();
    this.gradientFillIds = new HashMap<>();
    indexExistingGradientFills();
  }

  /**
   * Returns the current cell style with the local-date number format merged on top.
   *
   * <p>This preserves any existing fill, border, font, alignment, or wrap state already present on
   * the cell.
   */
  public CellStyle localDateStyle(Cell cell) {
    return mergedStyle(cell, LOCAL_DATE_STYLE_PATCH);
  }

  /**
   * Returns the current cell style with the local-date-time number format merged on top.
   *
   * <p>This preserves any existing fill, border, font, alignment, or wrap state already present on
   * the cell.
   */
  public CellStyle localDateTimeStyle(Cell cell) {
    return mergedStyle(cell, LOCAL_DATE_TIME_STYLE_PATCH);
  }

  /** Returns the workbook's default cell style (index 0). */
  public CellStyle defaultStyle() {
    return defaultStyleRecord();
  }

  /**
   * Resolves the current cell style and merges the provided style patch on top of it, returning a
   * cached or newly created {@link CellStyle}.
   */
  public CellStyle mergedStyle(Cell cell, CellStylePatch stylePatch) {
    return styleFor(styleRecord(cell), stylePatch);
  }

  /** Captures a read-only snapshot of the effective style applied to the given cell. */
  public CellStyleReport snapshot(Cell cell) {
    XSSFCellStyle style = styleRecord(cell);
    return snapshot(style);
  }

  /** Captures a read-only snapshot of the workbook's default cell style. */
  public CellStyleReport defaultSnapshot() {
    return snapshot(defaultStyleRecord());
  }

  private XSSFCellStyle defaultStyleRecord() {
    return workbook.getCellStyleAt(0);
  }

  private XSSFCellStyle styleRecord(Cell cell) {
    return (XSSFCellStyle) cell.getCellStyle();
  }

  private XSSFCellStyle styleFor(XSSFCellStyle baseStyle, CellStylePatch stylePatch) {
    CellStylePatch normalized = normalizedStylePatch(stylePatch);
    return cellStyles.computeIfAbsent(
        new MergedCellStyleKey(baseStyle.getIndex(), normalized),
        key -> createMergedStyle(baseStyle, normalized));
  }

  private static CellStylePatch normalizedStylePatch(CellStylePatch patch) {
    if (patch.font().isEmpty()) {
      return patch;
    }
    CellFontPatch font = patch.font().orElseThrow();
    if (font.fontHeight().isEmpty()
        || font.fontHeight().orElseThrow() instanceof FontHeightSpec.Twips) {
      return patch;
    }
    CellFontPatch normalizedFont =
        new CellFontPatch(
            font.bold(),
            font.italic(),
            font.fontName(),
            Optional.of(new FontHeightSpec.Twips(font.fontHeight().orElseThrow().twips())),
            font.fontColor(),
            font.underline(),
            font.strikeout());
    return new CellStylePatch(
        patch.numberFormat(),
        patch.alignment(),
        Optional.of(normalizedFont),
        patch.fill(),
        patch.border(),
        patch.protection());
  }

  private XSSFCellStyle createMergedStyle(XSSFCellStyle baseStyle, CellStylePatch stylePatch) {
    requireCellStyleCapacity(workbook); // LIM-011
    XSSFCellStyle cellStyle = workbook.createCellStyle();
    cellStyle.cloneStyleFrom(baseStyle);

    stylePatch
        .numberFormat()
        .ifPresent(numberFormat -> cellStyle.setDataFormat(dataFormat.getFormat(numberFormat)));
    applyAlignmentPatch(cellStyle, stylePatch.alignment());
    applyFillPatch(cellStyle, stylePatch.fill());
    stylePatch.border().ifPresent(border -> applyBorderPatch(cellStyle, border));
    stylePatch.protection().ifPresent(protection -> applyProtectionPatch(cellStyle, protection));
    stylePatch
        .font()
        .ifPresent(fontPatch -> cellStyle.setFont(fontFor(baseStyle.getFont(), fontPatch)));
    return cellStyle;
  }

  private void applyAlignmentPatch(
      XSSFCellStyle cellStyle, Optional<CellAlignmentPatch> alignmentPatch) {
    if (alignmentPatch.isEmpty()) {
      return;
    }
    CellAlignmentPatch patch = alignmentPatch.orElseThrow();
    patch.wrapText().ifPresent(cellStyle::setWrapText);
    patch.horizontalAlignment().ifPresent(value -> cellStyle.setAlignment(toPoi(value)));
    patch.verticalAlignment().ifPresent(value -> cellStyle.setVerticalAlignment(toPoi(value)));
    patch.textRotation().ifPresent(value -> cellStyle.setRotation(value.shortValue()));
    patch.indentation().ifPresent(value -> cellStyle.setIndention(value.shortValue()));
  }

  private void applyFillPatch(XSSFCellStyle cellStyle, Optional<CellFillSpec> fillPatch) {
    if (fillPatch.isEmpty()) {
      return;
    }
    switch (fillPatch.orElseThrow()) {
      case CellFillSpec.Gradient gradient -> {
        applyGradientFillPatch(cellStyle, gradient.gradient());
        return;
      }
      case CellFillSpec.PatternOnly pattern -> {
        cellStyle.setFillPattern(toPoi(pattern.pattern()));
        if (pattern.pattern() == ExcelFillPattern.NONE) {
          clearFillColors(cellStyle);
          return;
        }
        if (pattern.pattern() == ExcelFillPattern.SOLID) {
          cellStyle.setFillBackgroundColor((XSSFColor) null);
        }
      }
      case CellFillSpec.PatternForeground pattern -> {
        cellStyle.setFillPattern(toPoi(pattern.pattern()));
        if (pattern.pattern() == ExcelFillPattern.SOLID) {
          cellStyle.setFillBackgroundColor((XSSFColor) null);
        }
        cellStyle.setFillForegroundColor(
            ColorAdapter.toXssfColor(workbook, pattern.foregroundColor()));
      }
      case CellFillSpec.PatternBackground pattern -> {
        cellStyle.setFillPattern(toPoi(pattern.pattern()));
        cellStyle.setFillBackgroundColor(
            ColorAdapter.toXssfColor(workbook, pattern.backgroundColor()));
      }
      case CellFillSpec.PatternForegroundBackground pattern -> {
        cellStyle.setFillPattern(toPoi(pattern.pattern()));
        cellStyle.setFillForegroundColor(
            ColorAdapter.toXssfColor(workbook, pattern.foregroundColor()));
        cellStyle.setFillBackgroundColor(
            ColorAdapter.toXssfColor(workbook, pattern.backgroundColor()));
      }
    }
  }

  private void applyProtectionPatch(XSSFCellStyle cellStyle, CellProtectionPatch protectionPatch) {
    protectionPatch.locked().ifPresent(cellStyle::setLocked);
    protectionPatch.hiddenFormula().ifPresent(cellStyle::setHidden);
  }

  private XSSFFont fontFor(XSSFFont baseFont, CellFontPatch fontPatch) {
    return fonts.computeIfAbsent(
        new MergedFontKey(baseFont.getIndex(), fontPatch),
        key -> createMergedFont(baseFont, fontPatch));
  }

  private XSSFFont createMergedFont(XSSFFont baseFont, CellFontPatch fontPatch) {
    XSSFFont font = workbook.createFont();
    font.getCTFont().set(baseFont.getCTFont());
    fontPatch.bold().ifPresent(font::setBold);
    fontPatch.italic().ifPresent(font::setItalic);
    fontPatch.fontName().ifPresent(font::setFontName);
    fontPatch
        .fontHeight()
        .ifPresent(fontHeight -> font.setFontHeight(fontHeight.points().doubleValue()));
    fontPatch.fontColor().ifPresent(color -> applyFontColor(font, color));
    fontPatch
        .underline()
        .ifPresent(
            underline -> font.setUnderline(underline ? FontUnderline.SINGLE : FontUnderline.NONE));
    fontPatch.strikeout().ifPresent(font::setStrikeout);
    return font;
  }

  private void applyFontColor(XSSFFont font, ColorSpec color) {
    CTFont ctFont = font.getCTFont();
    while (ctFont.sizeOfColorArray() > 1) {
      ctFont.removeColor(ctFont.sizeOfColorArray() - 1);
    }
    if (ctFont.sizeOfColorArray() == 0) {
      ctFont.addNewColor();
    }
    ctFont.getColorArray(0).set(ColorAdapter.toXssfColor(workbook, color).getCTColor());
  }

  private record MergedCellStyleKey(int baseStyleIndex, CellStylePatch stylePatch) {}

  private record MergedFontKey(int baseFontIndex, CellFontPatch fontPatch) {}

  private static final String DEFAULT_NUMBER_FORMAT = "General";

  public static CellFontReport snapshotFont(XSSFFont font) {
    return new CellFontReport(
        font.getBold(),
        font.getItalic(),
        font.getFontName(),
        FontHeightReport.fromTwips(font.getFontHeight()),
        ColorAdapter.snapshot(font.getXSSFColor()).orElse(null),
        font.getUnderline() != FontUnderline.NONE.getByteValue(),
        font.getStrikeout());
  }

  public static String resolveNumberFormat(String numberFormat) {
    return numberFormat == null || numberFormat.isBlank() ? DEFAULT_NUMBER_FORMAT : numberFormat;
  }

  public CellStyleReport snapshot(XSSFCellStyle style) {
    return new CellStyleReport(
        resolveNumberFormat(style.getDataFormatString()),
        new CellAlignmentReport(
            style.getWrapText(),
            fromPoi(style.getAlignment()),
            fromPoi(style.getVerticalAlignment()),
            style.getRotation(),
            style.getIndention()),
        snapshotFont(style.getFont()),
        fillSnapshot(style),
        new CellBorderReport(
            borderSideSnapshot(style.getBorderTop(), style.getTopBorderXSSFColor()),
            borderSideSnapshot(style.getBorderRight(), style.getRightBorderXSSFColor()),
            borderSideSnapshot(style.getBorderBottom(), style.getBottomBorderXSSFColor()),
            borderSideSnapshot(style.getBorderLeft(), style.getLeftBorderXSSFColor())),
        new CellProtectionReport(style.getLocked(), style.getHidden()));
  }

  public CellGradientFillReport gradientFillSnapshot(CTGradientFill fill) {
    Double left = fill.isSetLeft() ? fill.getLeft() : null;
    Double right = fill.isSetRight() ? fill.getRight() : null;
    Double top = fill.isSetTop() ? fill.getTop() : null;
    Double bottom = fill.isSetBottom() ? fill.getBottom() : null;
    String type =
        ExcelGradientFillGeometry.effectiveType(
            fill.isSetType() ? fill.getType().toString() : null, left, right, top, bottom);
    java.util.List<CellGradientStopReport> stops =
        java.util.Arrays.stream(fill.getStopArray()).map(this::gradientStopSnapshot).toList();
    return "PATH".equals(type)
        ? CellGradientFillReport.path(left, right, top, bottom, stops)
        : CellGradientFillReport.linear(fill.isSetDegree() ? fill.getDegree() : null, stops);
  }

  private CellFillReport fillSnapshot(XSSFCellStyle style) {
    XSSFCellFill fill = fill(style);
    if (fill.getCTFill().isSetGradientFill()) {
      return CellFillReport.gradient(gradientFillSnapshot(fill.getCTFill().getGradientFill()));
    }
    var pattern = fromPoi(style.getFillPattern());
    if (pattern == dev.erst.gridgrind.core.vocabulary.ExcelFillPattern.NONE) {
      return CellFillReport.pattern(pattern);
    }
    Optional<CellColorReport> foreground =
        ColorAdapter.snapshot(style.getFillForegroundColorColor());
    Optional<CellColorReport> background =
        pattern == dev.erst.gridgrind.core.vocabulary.ExcelFillPattern.SOLID
            ? Optional.empty()
            : ColorAdapter.snapshot(style.getFillBackgroundColorColor());
    if (foreground.isPresent() && background.isPresent()) {
      return CellFillReport.patternColors(
          pattern, foreground.orElseThrow(), background.orElseThrow());
    }
    if (foreground.isPresent()) {
      return CellFillReport.patternForeground(pattern, foreground.orElseThrow());
    }
    if (background.isPresent()) {
      return CellFillReport.patternBackground(pattern, background.orElseThrow());
    }
    return CellFillReport.pattern(pattern);
  }

  private XSSFCellFill fill(XSSFCellStyle style) {
    long fillId = style.getCoreXf().getFillId();
    return workbook.getStylesSource().getFillAt((int) fillId);
  }

  private CellGradientStopReport gradientStopSnapshot(CTGradientStop stop) {
    CellColorReport color =
        ColorAdapter.snapshot(workbook, stop.getColor())
            .orElseThrow(
                () ->
                    new IllegalStateException(
                        "Gradient stop at position "
                            + stop.getPosition()
                            + " is missing its color definition."));
    return new CellGradientStopReport(stop.getPosition(), color);
  }

  private static CellBorderSideReport borderSideSnapshot(
      BorderStyle borderStyle, XSSFColor borderColor) {
    var style = fromPoi(borderStyle);
    if (style == dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle.NONE) {
      return new CellBorderSideReport.None();
    }
    return ColorAdapter.snapshot(borderColor)
        .<CellBorderSideReport>map(color -> new CellBorderSideReport.Colored(style, color))
        .orElseGet(() -> new CellBorderSideReport.DefaultColor(style));
  }

  /** Builds the POI rich-text payload for one authored GridGrind rich-text value. */
  public static XSSFRichTextString toPoiRichText(XSSFWorkbook workbook, ResolvedRichText richText) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(richText, "richText must not be null");

    CellTextLimits.requireSupportedLength(richText.plainText(), "richText.text"); // LIM-010
    XSSFRichTextString poiRichText = new XSSFRichTextString();
    for (ResolvedRichTextRun run : richText.runs()) {
      poiRichText.append(
          run.text(), run.font().map(font -> fontPatch(workbook, font)).orElse(null));
    }
    return poiRichText;
  }

  /** Returns factual rich-text runs, or empty when the cell stores only a scalar plain string. */
  public static Optional<List<RichTextRunReport>> snapshot(
      XSSFWorkbook workbook, XSSFRichTextString richText, CellFontReport baseFont) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(richText, "richText must not be null");
    Objects.requireNonNull(baseFont, "baseFont must not be null");

    CTRElt[] storedRuns = richText.getCTRst().getRArray();
    if (storedRuns.length == 0) {
      return Optional.empty();
    }

    List<RichTextRunReport> runs = new ArrayList<>(storedRuns.length);
    for (CTRElt run : storedRuns) {
      runs.add(new RichTextRunReport(run.getT(), merge(baseFont, runFontPatch(workbook, run))));
    }
    return Optional.of(List.copyOf(runs));
  }

  private static XSSFFont fontPatch(XSSFWorkbook workbook, CellFontPatch fontPatch) {
    CTFont font = CTFont.Factory.newInstance();
    setBooleanProperty(fontPatch.bold(), font::addNewB);
    setBooleanProperty(fontPatch.italic(), font::addNewI);
    if (fontPatch.fontName().isPresent()) {
      font.addNewName().setVal(fontPatch.fontName().orElseThrow());
    }
    if (fontPatch.fontHeight().isPresent()) {
      font.addNewSz().setVal(fontPatch.fontHeight().orElseThrow().points().doubleValue());
    }
    if (fontPatch.fontColor().isPresent()) {
      font.addNewColor()
          .set(
              ColorAdapter.toXssfColor(workbook, fontPatch.fontColor().orElseThrow()).getCTColor());
    }
    if (fontPatch.underline().isPresent()) {
      CTUnderlineProperty underline = font.addNewU();
      underline.setVal(
          fontPatch.underline().orElseThrow() ? STUnderlineValues.SINGLE : STUnderlineValues.NONE);
    }
    setBooleanProperty(fontPatch.strikeout(), font::addNewStrike);
    return new XSSFFont(font);
  }

  private static CellFontReport merge(CellFontReport baseFont, RunFontPatch fontPatch) {
    if (fontPatch.isEmpty()) {
      return baseFont;
    }
    return new CellFontReport(
        fontPatch.bold().orElse(baseFont.bold()),
        fontPatch.italic().orElse(baseFont.italic()),
        fontPatch.fontName().orElse(baseFont.fontName()),
        fontPatch.fontHeight().orElse(baseFont.fontHeight()),
        fontPatch.fontColor().isPresent()
            ? fontPatch.fontColor().orElseThrow()
            : baseFont.fontColor(),
        fontPatch.underline().orElse(baseFont.underline()),
        fontPatch.strikeout().orElse(baseFont.strikeout()));
  }

  private static RunFontPatch runFontPatch(XSSFWorkbook workbook, CTRElt run) {
    CTRPrElt properties = run.getRPr();
    if (properties == null) {
      return RunFontPatch.empty();
    }

    return new RunFontPatch(
        readBold(properties),
        readItalic(properties),
        readFontName(properties),
        readFontHeight(properties),
        readFontColor(workbook, properties),
        readUnderline(properties),
        readStrikeout(properties));
  }

  private static Optional<Boolean> readBold(CTRPrElt properties) {
    return booleanProperty(properties.sizeOfBArray() > 0 ? properties.getBArray(0) : null);
  }

  private static Optional<Boolean> readItalic(CTRPrElt properties) {
    return booleanProperty(properties.sizeOfIArray() > 0 ? properties.getIArray(0) : null);
  }

  private static Optional<String> readFontName(CTRPrElt properties) {
    return properties.sizeOfRFontArray() > 0
        ? Optional.of(properties.getRFontArray(0).getVal())
        : Optional.empty();
  }

  private static Optional<FontHeightReport> readFontHeight(CTRPrElt properties) {
    if (properties.sizeOfSzArray() == 0) {
      return Optional.empty();
    }
    return Optional.of(
        FontHeightReport.fromPoints(
            java.math.BigDecimal.valueOf(properties.getSzArray(0).getVal())));
  }

  private static Optional<CellColorReport> readFontColor(
      XSSFWorkbook workbook, CTRPrElt properties) {
    return properties.sizeOfColorArray() > 0
        ? ColorAdapter.snapshot(workbook, properties.getColorArray(0))
        : Optional.empty();
  }

  private static Optional<Boolean> readUnderline(CTRPrElt properties) {
    return properties.sizeOfUArray() > 0
        ? Optional.of(underline(properties.getUArray(0)))
        : Optional.empty();
  }

  private static Optional<Boolean> readStrikeout(CTRPrElt properties) {
    return booleanProperty(
        properties.sizeOfStrikeArray() > 0 ? properties.getStrikeArray(0) : null);
  }

  private static Optional<Boolean> booleanProperty(@Nullable CTBooleanProperty property) {
    return property == null
        ? Optional.empty()
        : Optional.of(!property.isSetVal() || property.getVal());
  }

  private static boolean underline(CTUnderlineProperty property) {
    Objects.requireNonNull(property, "property must not be null");
    return !property.isSetVal() || property.getVal() != STUnderlineValues.NONE;
  }

  private static void setBooleanProperty(
      Optional<Boolean> value, Supplier<CTBooleanProperty> supplier) {
    Objects.requireNonNull(value, "value must not be null");
    value.ifPresent(booleanValue -> supplier.get().setVal(booleanValue));
  }

  /** Partial font override captured from one rich-text run. */
  private record RunFontPatch(
      Optional<Boolean> bold,
      Optional<Boolean> italic,
      Optional<String> fontName,
      Optional<FontHeightReport> fontHeight,
      Optional<CellColorReport> fontColor,
      Optional<Boolean> underline,
      Optional<Boolean> strikeout) {
    private static RunFontPatch empty() {
      return new RunFontPatch(
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          Optional.empty(),
          Optional.empty());
    }

    private boolean isEmpty() {
      return bold.isEmpty()
          && italic.isEmpty()
          && fontName.isEmpty()
          && fontHeight.isEmpty()
          && fontColor.isEmpty()
          && underline.isEmpty()
          && strikeout.isEmpty();
    }
  }

  public void applyBorderPatch(XSSFCellStyle cellStyle, CellBorderPatch border) {
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.top()),
        cellStyle::setBorderTop,
        cellStyle::setTopBorderColor);
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.right()),
        cellStyle::setBorderRight,
        cellStyle::setRightBorderColor);
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.bottom()),
        cellStyle::setBorderBottom,
        cellStyle::setBottomBorderColor);
    applyBorderSidePatch(
        mergedBorderSide(border.all(), border.left()),
        cellStyle::setBorderLeft,
        cellStyle::setLeftBorderColor);
  }

  private Optional<BorderSidePatch> mergedBorderSide(
      Optional<BorderSidePatch> defaultSide, Optional<BorderSidePatch> explicitSide) {
    return effectiveBorderSide(
        mergedBorderStyle(defaultSide, explicitSide), mergedBorderColor(defaultSide, explicitSide));
  }

  private Optional<ExcelBorderStyle> mergedBorderStyle(
      Optional<BorderSidePatch> defaultSide, Optional<BorderSidePatch> explicitSide) {
    Objects.requireNonNull(defaultSide, "defaultSide must not be null");
    Objects.requireNonNull(explicitSide, "explicitSide must not be null");
    Optional<ExcelBorderStyle> explicitStyle = explicitSide.flatMap(BorderSidePatch::style);
    return explicitStyle.isPresent() ? explicitStyle : defaultSide.flatMap(BorderSidePatch::style);
  }

  private Optional<ColorSpec> mergedBorderColor(
      Optional<BorderSidePatch> defaultSide, Optional<BorderSidePatch> explicitSide) {
    Objects.requireNonNull(defaultSide, "defaultSide must not be null");
    Objects.requireNonNull(explicitSide, "explicitSide must not be null");
    if (explicitSide.isPresent() && explicitSide.orElseThrow().color().isPresent()) {
      return explicitSide.orElseThrow().color();
    }
    return defaultSide.isEmpty() || defaultSide.orElseThrow().color().isEmpty()
        ? Optional.empty()
        : defaultSide.orElseThrow().color();
  }

  private Optional<BorderSidePatch> effectiveBorderSide(
      Optional<ExcelBorderStyle> style, Optional<ColorSpec> color) {
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(color, "color must not be null");
    if (style.isEmpty() && color.isEmpty()) {
      return Optional.empty();
    }
    if (color.isPresent() && (style.isEmpty() || style.orElseThrow() == ExcelBorderStyle.NONE)) {
      throw new IllegalArgumentException("border side color requires an effective border style");
    }
    return Optional.of(new BorderSidePatch(style, color));
  }

  private void applyBorderSidePatch(
      Optional<BorderSidePatch> sidePatch,
      Consumer<BorderStyle> styleSetter,
      Consumer<XSSFColor> colorSetter) {
    Objects.requireNonNull(sidePatch, "sidePatch must not be null");
    if (sidePatch.isEmpty()) {
      return;
    }
    BorderSidePatch resolved = sidePatch.orElseThrow();
    styleSetter.accept(toPoi(resolved.style().orElseThrow()));
    if (resolved.style().orElseThrow() == ExcelBorderStyle.NONE) {
      return;
    }
    if (resolved.color().isPresent()) {
      colorSetter.accept(ColorAdapter.toXssfColor(workbook, resolved.color().orElseThrow()));
    }
  }

  public void indexExistingGradientFills() {
    List<XSSFCellFill> fills = fillsList();
    for (int fillId = 0; fillId < fills.size(); fillId++) {
      XSSFCellFill fill = fills.get(fillId);
      if (fill.getCTFill().isSetGradientFill()) {
        gradientFillIds.putIfAbsent(fill.getCTFill().xmlText(), fillId);
      }
    }
  }

  public void applyGradientFillPatch(XSSFCellStyle cellStyle, CellGradientFillSpec gradientPatch) {
    CTFill gradientFill = CTFill.Factory.newInstance();
    CTGradientFill gradient = gradientFill.addNewGradientFill();
    switch (gradientPatch) {
      case CellGradientFillSpec.Path path -> {
        gradient.setType(
            org.openxmlformats.schemas.spreadsheetml.x2006.main.STGradientType.Enum.forString(
                "path"));
        path.left().ifPresent(gradient::setLeft);
        path.right().ifPresent(gradient::setRight);
        path.top().ifPresent(gradient::setTop);
        path.bottom().ifPresent(gradient::setBottom);
      }
      case CellGradientFillSpec.Linear linear -> linear.degree().ifPresent(gradient::setDegree);
    }
    for (CellGradientStopSpec stop : gradientPatch.stops()) {
      CTGradientStop ctStop = gradient.addNewStop();
      ctStop.setPosition(stop.position());
      ctStop.addNewColor().set(ColorAdapter.toXssfColor(workbook, stop.color()).getCTColor());
    }
    int gradientFillId = gradientFillId(gradientFill);
    cellStyle.getCoreXf().setApplyFill(true);
    cellStyle.getCoreXf().setFillId(gradientFillId);
  }

  public void clearFillColors(XSSFCellStyle cellStyle) {
    cellStyle.setFillForegroundColor((XSSFColor) null);
    cellStyle.setFillBackgroundColor((XSSFColor) null);
  }

  private int gradientFillId(CTFill gradientFill) {
    String key = gradientFill.xmlText();
    Integer existingId = gradientFillIds.get(key);
    if (existingId != null) {
      return existingId;
    }
    int fillId =
        appendFill(new XSSFCellFill(gradientFill, workbook.getStylesSource().getIndexedColors()));
    gradientFillIds.put(key, fillId);
    return fillId;
  }

  private int appendFill(XSSFCellFill fill) {
    return fillRegistryAccess.appendFill(workbook.getStylesSource(), fill);
  }

  private List<XSSFCellFill> fillsList() {
    return fillRegistryAccess.fills(workbook.getStylesSource());
  }

  private static final Map<FillPatternType, ExcelFillPattern> FROM_POI_FILL_PATTERN =
      EnumMappings.exactEnumMap(
          FillPatternType.class,
          "Excel cell fill-pattern mapping",
          Map.ofEntries(
              Map.entry(FillPatternType.NO_FILL, ExcelFillPattern.NONE),
              Map.entry(FillPatternType.SOLID_FOREGROUND, ExcelFillPattern.SOLID),
              Map.entry(FillPatternType.FINE_DOTS, ExcelFillPattern.FINE_DOTS),
              Map.entry(FillPatternType.ALT_BARS, ExcelFillPattern.ALT_BARS),
              Map.entry(FillPatternType.SPARSE_DOTS, ExcelFillPattern.SPARSE_DOTS),
              Map.entry(FillPatternType.THICK_HORZ_BANDS, ExcelFillPattern.THICK_HORIZONTAL_BANDS),
              Map.entry(FillPatternType.THICK_VERT_BANDS, ExcelFillPattern.THICK_VERTICAL_BANDS),
              Map.entry(
                  FillPatternType.THICK_BACKWARD_DIAG, ExcelFillPattern.THICK_BACKWARD_DIAGONAL),
              Map.entry(
                  FillPatternType.THICK_FORWARD_DIAG, ExcelFillPattern.THICK_FORWARD_DIAGONAL),
              Map.entry(FillPatternType.BIG_SPOTS, ExcelFillPattern.BIG_SPOTS),
              Map.entry(FillPatternType.BRICKS, ExcelFillPattern.BRICKS),
              Map.entry(FillPatternType.THIN_HORZ_BANDS, ExcelFillPattern.THIN_HORIZONTAL_BANDS),
              Map.entry(FillPatternType.THIN_VERT_BANDS, ExcelFillPattern.THIN_VERTICAL_BANDS),
              Map.entry(
                  FillPatternType.THIN_BACKWARD_DIAG, ExcelFillPattern.THIN_BACKWARD_DIAGONAL),
              Map.entry(FillPatternType.THIN_FORWARD_DIAG, ExcelFillPattern.THIN_FORWARD_DIAGONAL),
              Map.entry(FillPatternType.SQUARES, ExcelFillPattern.SQUARES),
              Map.entry(FillPatternType.DIAMONDS, ExcelFillPattern.DIAMONDS),
              Map.entry(FillPatternType.LESS_DOTS, ExcelFillPattern.LESS_DOTS),
              Map.entry(FillPatternType.LEAST_DOTS, ExcelFillPattern.LEAST_DOTS)));

  private static final Map<ExcelFillPattern, FillPatternType> TO_POI_FILL_PATTERN =
      EnumMappings.reverseExactEnumMap(
          ExcelFillPattern.class, "Apache POI fill-pattern mapping", FROM_POI_FILL_PATTERN);

  public static ExcelHorizontalAlignment fromPoi(HorizontalAlignment alignment) {
    return ExcelHorizontalAlignment.valueOf(alignment.name());
  }

  public static ExcelVerticalAlignment fromPoi(VerticalAlignment alignment) {
    return ExcelVerticalAlignment.valueOf(alignment.name());
  }

  public static ExcelBorderStyle fromPoi(BorderStyle borderStyle) {
    return ExcelBorderStyle.valueOf(borderStyle.name());
  }

  public static HorizontalAlignment toPoi(ExcelHorizontalAlignment alignment) {
    return HorizontalAlignment.valueOf(alignment.name());
  }

  public static VerticalAlignment toPoi(ExcelVerticalAlignment alignment) {
    return VerticalAlignment.valueOf(alignment.name());
  }

  public static BorderStyle toPoi(ExcelBorderStyle borderStyle) {
    return BorderStyle.valueOf(borderStyle.name());
  }

  public static ExcelFillPattern fromPoi(FillPatternType pattern) {
    return EnumMappings.requireMappedValue(FROM_POI_FILL_PATTERN, pattern, "POI fill pattern");
  }

  public static FillPatternType toPoi(ExcelFillPattern pattern) {
    return EnumMappings.requireMappedValue(TO_POI_FILL_PATTERN, pattern, "GridGrind fill pattern");
  }

  public static final int MAX_CELL_STYLES =
      SpreadsheetVersion.EXCEL2007.getMaxCellStyles(); // LIM-011

  public static void requireCellStyleCapacity(XSSFWorkbook workbook) {
    requireCellStyleCapacity(workbook.getNumCellStyles());
  }

  public static void requireCellStyleCapacity(int existingStyleCount) {
    if (existingStyleCount >= MAX_CELL_STYLES) {
      throw new IllegalArgumentException(
          "workbook cannot create more than "
              + MAX_CELL_STYLES
              + " cell styles (Excel/POI workbook style limit)");
    }
  }
}
