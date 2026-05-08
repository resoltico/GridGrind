package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Supplier;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRPrElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues;

/** Project-owned rich-text seam used to author and snapshot string-cell runs. */
final class ExcelRichTextSupport {
  private ExcelRichTextSupport() {}

  /** Builds the POI rich-text payload for one authored GridGrind rich-text value. */
  static XSSFRichTextString toPoiRichText(XSSFWorkbook workbook, ExcelRichText richText) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(richText, "richText must not be null");

    ExcelCellTextLimits.requireSupportedLength(richText.plainText(), "richText.text"); // LIM-010
    XSSFRichTextString poiRichText = new XSSFRichTextString();
    for (ExcelRichTextRun run : richText.runs()) {
      poiRichText.append(
          run.text(), run.font().map(font -> fontPatch(workbook, font)).orElse(null));
    }
    return poiRichText;
  }

  /** Returns factual rich-text runs, or empty when the cell stores only a scalar plain string. */
  static Optional<ExcelRichTextSnapshot> snapshot(
      XSSFWorkbook workbook, XSSFRichTextString richText, ExcelCellFontSnapshot baseFont) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(richText, "richText must not be null");
    Objects.requireNonNull(baseFont, "baseFont must not be null");

    CTRElt[] runs = richText.getCTRst().getRArray();
    if (runs.length == 0) {
      return Optional.empty();
    }

    List<ExcelRichTextRunSnapshot> snapshots = new ArrayList<>(runs.length);
    for (CTRElt run : runs) {
      snapshots.add(
          new ExcelRichTextRunSnapshot(run.getT(), merge(baseFont, runFontPatch(workbook, run))));
    }
    return Optional.of(new ExcelRichTextSnapshot(snapshots));
  }

  private static XSSFFont fontPatch(XSSFWorkbook workbook, ExcelCellFont fontPatch) {
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
              ExcelColorSupport.toXssfColor(workbook, fontPatch.fontColor().orElseThrow())
                  .getCTColor());
    }
    if (fontPatch.underline().isPresent()) {
      CTUnderlineProperty underline = font.addNewU();
      underline.setVal(
          fontPatch.underline().orElseThrow() ? STUnderlineValues.SINGLE : STUnderlineValues.NONE);
    }
    setBooleanProperty(fontPatch.strikeout(), font::addNewStrike);
    return new XSSFFont(font);
  }

  private static ExcelCellFontSnapshot merge(
      ExcelCellFontSnapshot baseFont, RunFontPatch fontPatch) {
    if (fontPatch.isEmpty()) {
      return baseFont;
    }
    return new ExcelCellFontSnapshot(
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

  private static Optional<ExcelFontHeight> readFontHeight(CTRPrElt properties) {
    if (properties.sizeOfSzArray() == 0) {
      return Optional.empty();
    }
    return Optional.of(
        ExcelFontHeight.fromPoints(
            java.math.BigDecimal.valueOf(properties.getSzArray(0).getVal())));
  }

  private static Optional<ExcelColorSnapshot> readFontColor(
      XSSFWorkbook workbook, CTRPrElt properties) {
    return properties.sizeOfColorArray() > 0
        ? ExcelColorSnapshotSupport.snapshot(workbook, properties.getColorArray(0))
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
      Optional<ExcelFontHeight> fontHeight,
      Optional<ExcelColorSnapshot> fontColor,
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
}
