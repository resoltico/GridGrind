package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.xssf.model.StylesTable;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBooleanProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorderPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFont;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFontSize;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTNumFmt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTUnderlineProperty;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnderlineValues;

/** Materializes authored conditional-formatting styles into OOXML differential-style payloads. */
final class ExcelConditionalFormattingStyleWriteSupport {
  private ExcelConditionalFormattingStyleWriteSupport() {}

  static void applyStyle(StylesTable stylesTable, CTDxf dxf, ExcelDifferentialStyle style) {
    applyNumberFormat(stylesTable, dxf, style.numberFormat().orElse(null));
    applyFont(dxf, style);
    applyFill(dxf, style.fillColor().orElse(null));
    applyBorder(dxf, style.border().orElse(null));
  }

  private static void applyNumberFormat(
      StylesTable stylesTable, CTDxf dxf, @Nullable String numberFormat) {
    if (numberFormat == null) {
      return;
    }
    int formatId = stylesTable.putNumberFormat(numberFormat);
    CTNumFmt numFmt = dxf.addNewNumFmt();
    numFmt.setNumFmtId(formatId);
    numFmt.setFormatCode(numberFormat);
  }

  private static void applyFont(CTDxf dxf, ExcelDifferentialStyle style) {
    if (java.util.stream.Stream.of(
            style.bold().orElse(null),
            style.italic().orElse(null),
            style.fontHeight().orElse(null),
            style.fontColor().orElse(null),
            style.underline().orElse(null),
            style.strikeout().orElse(null))
        .allMatch(Objects::isNull)) {
      return;
    }

    CTFont font = dxf.addNewFont();
    setBooleanProperty(style.bold().orElse(null), font::addNewB);
    setBooleanProperty(style.italic().orElse(null), font::addNewI);
    setBooleanProperty(style.strikeout().orElse(null), font::addNewStrike);
    if (style.fontHeight().isPresent()) {
      CTFontSize fontSize = font.addNewSz();
      fontSize.setVal(style.fontHeight().orElseThrow().points().doubleValue());
    }
    if (style.fontColor().isPresent()) {
      ExcelConditionalFormattingColorSupport.setColor(
          font.addNewColor(), style.fontColor().orElseThrow());
    }
    if (style.underline().isPresent()) {
      CTUnderlineProperty underlineProperty = font.addNewU();
      underlineProperty.setVal(
          style.underline().orElseThrow() ? STUnderlineValues.SINGLE : STUnderlineValues.NONE);
    }
  }

  private static void applyFill(CTDxf dxf, @Nullable String fillColor) {
    if (fillColor == null) {
      return;
    }
    CTFill fill = dxf.addNewFill();
    CTPatternFill patternFill = fill.addNewPatternFill();
    patternFill.setPatternType(STPatternType.SOLID);
    ExcelConditionalFormattingColorSupport.setColor(patternFill.addNewFgColor(), fillColor);
  }

  private static void applyBorder(CTDxf dxf, @Nullable ExcelDifferentialBorder border) {
    if (border == null) {
      return;
    }
    CTBorder ctBorder = dxf.addNewBorder();
    applyBorderSide(ctBorder::addNewTop, resolvedSide(border.all(), border.top()));
    applyBorderSide(ctBorder::addNewRight, resolvedSide(border.all(), border.right()));
    applyBorderSide(ctBorder::addNewBottom, resolvedSide(border.all(), border.bottom()));
    applyBorderSide(ctBorder::addNewLeft, resolvedSide(border.all(), border.left()));
  }

  private static void applyBorderSide(
      java.util.function.Supplier<CTBorderPr> borderPrFactory,
      @Nullable ExcelDifferentialBorderSide side) {
    Objects.requireNonNull(borderPrFactory, "borderPrFactory must not be null");
    if (side == null) {
      return;
    }
    CTBorderPr borderPr = borderPrFactory.get();
    borderPr.setStyle(ExcelConditionalFormattingBorderStyleSupport.toCtBorderStyle(side.style()));
    if (side.color() != null) {
      ExcelConditionalFormattingColorSupport.setColor(borderPr.addNewColor(), side.color());
    }
  }

  private static @Nullable ExcelDifferentialBorderSide resolvedSide(
      @Nullable ExcelDifferentialBorderSide defaultSide,
      @Nullable ExcelDifferentialBorderSide explicitSide) {
    return explicitSide == null ? defaultSide : explicitSide;
  }

  private static void setBooleanProperty(
      @Nullable Boolean value, java.util.function.Supplier<CTBooleanProperty> propertyFactory) {
    if (value == null) {
      return;
    }
    CTBooleanProperty property = propertyFactory.get();
    property.setVal(value);
  }
}
