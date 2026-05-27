package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Map;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle;

/** Converts conditional-formatting border styles between GridGrind and OOXML forms. */
final class ExcelConditionalFormattingBorderStyleSupport {
  private static final Map<ExcelBorderStyle, STBorderStyle.Enum> TO_CT_BORDER_STYLE =
      ExcelEnumMappingSupport.exactEnumMap(
          ExcelBorderStyle.class,
          "Conditional-formatting CT border-style mapping",
          Map.ofEntries(
              Map.entry(ExcelBorderStyle.NONE, STBorderStyle.NONE),
              Map.entry(ExcelBorderStyle.THIN, STBorderStyle.THIN),
              Map.entry(ExcelBorderStyle.MEDIUM, STBorderStyle.MEDIUM),
              Map.entry(ExcelBorderStyle.DASHED, STBorderStyle.DASHED),
              Map.entry(ExcelBorderStyle.DOTTED, STBorderStyle.DOTTED),
              Map.entry(ExcelBorderStyle.THICK, STBorderStyle.THICK),
              Map.entry(ExcelBorderStyle.DOUBLE, STBorderStyle.DOUBLE),
              Map.entry(ExcelBorderStyle.HAIR, STBorderStyle.HAIR),
              Map.entry(ExcelBorderStyle.MEDIUM_DASHED, STBorderStyle.MEDIUM_DASHED),
              Map.entry(ExcelBorderStyle.DASH_DOT, STBorderStyle.DASH_DOT),
              Map.entry(ExcelBorderStyle.MEDIUM_DASH_DOT, STBorderStyle.MEDIUM_DASH_DOT),
              Map.entry(ExcelBorderStyle.DASH_DOT_DOT, STBorderStyle.DASH_DOT_DOT),
              Map.entry(ExcelBorderStyle.MEDIUM_DASH_DOT_DOT, STBorderStyle.MEDIUM_DASH_DOT_DOT),
              Map.entry(ExcelBorderStyle.SLANTED_DASH_DOT, STBorderStyle.SLANT_DASH_DOT)));

  private static final Map<Integer, ExcelBorderStyle> FROM_CT_BORDER_STYLE =
      Map.ofEntries(
          Map.entry(STBorderStyle.INT_NONE, ExcelBorderStyle.NONE),
          Map.entry(STBorderStyle.INT_THIN, ExcelBorderStyle.THIN),
          Map.entry(STBorderStyle.INT_MEDIUM, ExcelBorderStyle.MEDIUM),
          Map.entry(STBorderStyle.INT_DASHED, ExcelBorderStyle.DASHED),
          Map.entry(STBorderStyle.INT_DOTTED, ExcelBorderStyle.DOTTED),
          Map.entry(STBorderStyle.INT_THICK, ExcelBorderStyle.THICK),
          Map.entry(STBorderStyle.INT_DOUBLE, ExcelBorderStyle.DOUBLE),
          Map.entry(STBorderStyle.INT_HAIR, ExcelBorderStyle.HAIR),
          Map.entry(STBorderStyle.INT_MEDIUM_DASHED, ExcelBorderStyle.MEDIUM_DASHED),
          Map.entry(STBorderStyle.INT_DASH_DOT, ExcelBorderStyle.DASH_DOT),
          Map.entry(STBorderStyle.INT_MEDIUM_DASH_DOT, ExcelBorderStyle.MEDIUM_DASH_DOT),
          Map.entry(STBorderStyle.INT_DASH_DOT_DOT, ExcelBorderStyle.DASH_DOT_DOT),
          Map.entry(STBorderStyle.INT_MEDIUM_DASH_DOT_DOT, ExcelBorderStyle.MEDIUM_DASH_DOT_DOT),
          Map.entry(STBorderStyle.INT_SLANT_DASH_DOT, ExcelBorderStyle.SLANTED_DASH_DOT));

  private ExcelConditionalFormattingBorderStyleSupport() {}

  static STBorderStyle.Enum toCtBorderStyle(ExcelBorderStyle style) {
    return ExcelEnumMappingSupport.requireMappedValue(
        TO_CT_BORDER_STYLE, style, "GridGrind border style");
  }

  static ExcelBorderStyle fromCtBorderStyle(int styleCode) {
    ExcelBorderStyle resolved = FROM_CT_BORDER_STYLE.get(styleCode);
    if (resolved == null) {
      throw new IllegalArgumentException("Unsupported CT border style: " + styleCode);
    }
    return resolved;
  }
}
