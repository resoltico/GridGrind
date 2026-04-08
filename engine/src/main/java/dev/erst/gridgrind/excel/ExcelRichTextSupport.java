package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Supplier;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
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

    XSSFRichTextString poiRichText = new XSSFRichTextString();
    for (ExcelRichTextRun run : richText.runs()) {
      poiRichText.append(run.text(), run.font() == null ? null : fontPatch(workbook, run.font()));
    }
    return poiRichText;
  }

  /** Returns factual rich-text runs, or null when the cell stores only a scalar plain string. */
  static ExcelRichTextSnapshot snapshot(
      XSSFWorkbook workbook, XSSFRichTextString richText, ExcelCellFontSnapshot baseFont) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(richText, "richText must not be null");
    Objects.requireNonNull(baseFont, "baseFont must not be null");

    CTRElt[] runs = richText.getCTRst().getRArray();
    if (runs.length == 0) {
      return null;
    }

    List<ExcelRichTextRunSnapshot> snapshots = new ArrayList<>(runs.length);
    for (CTRElt run : runs) {
      snapshots.add(
          new ExcelRichTextRunSnapshot(run.getT(), merge(baseFont, runFontPatch(workbook, run))));
    }
    return new ExcelRichTextSnapshot(snapshots);
  }

  private static XSSFFont fontPatch(XSSFWorkbook workbook, ExcelCellFont fontPatch) {
    CTFont font = CTFont.Factory.newInstance();
    setBooleanProperty(fontPatch.bold(), font::addNewB);
    setBooleanProperty(fontPatch.italic(), font::addNewI);
    if (fontPatch.fontName() != null) {
      font.addNewName().setVal(fontPatch.fontName());
    }
    if (fontPatch.fontHeight() != null) {
      font.addNewSz().setVal(fontPatch.fontHeight().points().doubleValue());
    }
    if (fontPatch.fontColor() != null) {
      font.addNewColor()
          .setRgb(ExcelRgbColorSupport.toXssfColor(workbook, fontPatch.fontColor()).getARGB());
    }
    if (fontPatch.underline() != null) {
      CTUnderlineProperty underline = font.addNewU();
      underline.setVal(fontPatch.underline() ? STUnderlineValues.SINGLE : STUnderlineValues.NONE);
    }
    setBooleanProperty(fontPatch.strikeout(), font::addNewStrike);
    return new XSSFFont(font);
  }

  private static ExcelCellFontSnapshot merge(
      ExcelCellFontSnapshot baseFont, ExcelCellFont fontPatch) {
    if (fontPatch == null) {
      return baseFont;
    }
    return new ExcelCellFontSnapshot(
        fontPatch.bold() != null ? fontPatch.bold() : baseFont.bold(),
        fontPatch.italic() != null ? fontPatch.italic() : baseFont.italic(),
        fontPatch.fontName() != null ? fontPatch.fontName() : baseFont.fontName(),
        fontPatch.fontHeight() != null ? fontPatch.fontHeight() : baseFont.fontHeight(),
        fontPatch.fontColor() != null ? fontPatch.fontColor() : baseFont.fontColor(),
        fontPatch.underline() != null ? fontPatch.underline() : baseFont.underline(),
        fontPatch.strikeout() != null ? fontPatch.strikeout() : baseFont.strikeout());
  }

  private static ExcelCellFont runFontPatch(XSSFWorkbook workbook, CTRElt run) {
    CTRPrElt properties = run.getRPr();
    if (properties == null) {
      return null;
    }

    Boolean bold = readBold(properties);
    Boolean italic = readItalic(properties);
    String fontName = readFontName(properties);
    ExcelFontHeight fontHeight = readFontHeight(properties);
    String fontColor = readFontColor(workbook, properties);
    Boolean underline = readUnderline(properties);
    Boolean strikeout = readStrikeout(properties);
    if (allFontAttributesNull(
        bold, italic, fontName, fontHeight, fontColor, underline, strikeout)) {
      return null;
    }
    return new ExcelCellFont(bold, italic, fontName, fontHeight, fontColor, underline, strikeout);
  }

  private static Boolean readBold(CTRPrElt properties) {
    return booleanProperty(properties.sizeOfBArray() > 0 ? properties.getBArray(0) : null);
  }

  private static Boolean readItalic(CTRPrElt properties) {
    return booleanProperty(properties.sizeOfIArray() > 0 ? properties.getIArray(0) : null);
  }

  private static String readFontName(CTRPrElt properties) {
    return properties.sizeOfRFontArray() > 0 ? properties.getRFontArray(0).getVal() : null;
  }

  private static ExcelFontHeight readFontHeight(CTRPrElt properties) {
    if (properties.sizeOfSzArray() == 0) {
      return null;
    }
    return ExcelFontHeight.fromPoints(
        java.math.BigDecimal.valueOf(properties.getSzArray(0).getVal()));
  }

  private static String readFontColor(XSSFWorkbook workbook, CTRPrElt properties) {
    return properties.sizeOfColorArray() > 0
        ? toRgbHex(workbook, properties.getColorArray(0))
        : null;
  }

  private static Boolean readUnderline(CTRPrElt properties) {
    return properties.sizeOfUArray() > 0 ? underline(properties.getUArray(0)) : null;
  }

  private static Boolean readStrikeout(CTRPrElt properties) {
    return booleanProperty(
        properties.sizeOfStrikeArray() > 0 ? properties.getStrikeArray(0) : null);
  }

  private static boolean allFontAttributesNull(
      Boolean bold,
      Boolean italic,
      String fontName,
      ExcelFontHeight fontHeight,
      String fontColor,
      Boolean underline,
      Boolean strikeout) {
    return bold == null
        && italic == null
        && fontName == null
        && fontHeight == null
        && fontColor == null
        && underline == null
        && strikeout == null;
  }

  private static Boolean booleanProperty(CTBooleanProperty property) {
    return property == null ? null : !property.isSetVal() || property.getVal();
  }

  private static boolean underline(CTUnderlineProperty property) {
    Objects.requireNonNull(property, "property must not be null");
    return !property.isSetVal() || property.getVal() != STUnderlineValues.NONE;
  }

  private static void setBooleanProperty(Boolean value, Supplier<CTBooleanProperty> supplier) {
    if (value == null) {
      return;
    }
    supplier.get().setVal(value);
  }

  private static String toRgbHex(XSSFWorkbook workbook, CTColor color) {
    XSSFColor xssfColor = XSSFColor.from(color, workbook.getStylesSource().getIndexedColors());
    ThemesTable themes = workbook.getStylesSource().getTheme();
    if (themes != null) {
      themes.inheritFromThemeAsRequired(xssfColor);
    }
    return ExcelRgbColorSupport.toRgbHex(xssfColor);
  }
}
