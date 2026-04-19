package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;

/** Extracts stable style-attribute coverage labels from style patches. */
public final class StyleKindIntrospection {
  private StyleKindIntrospection() {}

  /** Returns attribute labels present on a protocol style patch. */
  public static Map<String, Long> styleKinds(CellStyleInput style) {
    Objects.requireNonNull(style, "style must not be null");
    LinkedHashMap<String, Long> kinds = new LinkedHashMap<>();
    increment(kinds, "number_format", style.numberFormat() != null);
    appendProtocolAlignmentKinds(kinds, style.alignment());
    appendProtocolFontKinds(kinds, style.font());
    appendProtocolFillKinds(kinds, style.fill());
    appendProtocolBorderKinds(kinds, style.border());
    appendProtocolProtectionKinds(kinds, style.protection());
    return Map.copyOf(kinds);
  }

  /** Returns attribute labels present on an engine style patch. */
  public static Map<String, Long> styleKinds(ExcelCellStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    LinkedHashMap<String, Long> kinds = new LinkedHashMap<>();
    increment(kinds, "number_format", style.numberFormat() != null);
    appendEngineAlignmentKinds(kinds, style);
    appendEngineFontKinds(kinds, style);
    appendEngineFillKinds(kinds, style);
    appendEngineBorderKinds(kinds, style.border());
    appendEngineProtectionKinds(kinds, style);
    return Map.copyOf(kinds);
  }

  private static void appendProtocolAlignmentKinds(
      Map<String, Long> kinds, CellAlignmentInput alignment) {
    increment(kinds, "alignment", alignment != null);
    if (alignment == null) {
      return;
    }
    increment(kinds, "wrap_text", alignment.wrapText() != null);
    increment(kinds, "horizontal_alignment", alignment.horizontalAlignment() != null);
    increment(kinds, "vertical_alignment", alignment.verticalAlignment() != null);
    increment(kinds, "text_rotation", alignment.textRotation() != null);
    increment(kinds, "indentation", alignment.indentation() != null);
  }

  private static void appendProtocolFontKinds(Map<String, Long> kinds, CellFontInput font) {
    increment(kinds, "font", font != null);
    if (font == null) {
      return;
    }
    increment(kinds, "bold", font.bold() != null);
    increment(kinds, "italic", font.italic() != null);
    increment(kinds, "font_name", font.fontName() != null);
    increment(kinds, "font_height", font.fontHeight() != null);
    increment(kinds, "font_height_points", font.fontHeight() instanceof FontHeightInput.Points);
    increment(kinds, "font_height_twips", font.fontHeight() instanceof FontHeightInput.Twips);
    increment(kinds, "font_color", font.fontColor() != null);
    increment(kinds, "underline", font.underline() != null);
    increment(kinds, "strikeout", font.strikeout() != null);
  }

  private static void appendProtocolFillKinds(Map<String, Long> kinds, CellFillInput fill) {
    increment(kinds, "fill", fill != null);
    if (fill == null) {
      return;
    }
    increment(kinds, "fill_pattern", fill.pattern() != null);
    increment(kinds, "fill_pattern_solid", fill.pattern() == ExcelFillPattern.SOLID);
    increment(kinds, "fill_patterned", isPatterned(fill.pattern()));
    increment(kinds, "fill_foreground_color", fill.foregroundColor() != null);
    increment(kinds, "fill_background_color", fill.backgroundColor() != null);
    increment(kinds, "fill_color", fill.foregroundColor() != null);
  }

  private static void appendProtocolBorderKinds(Map<String, Long> kinds, CellBorderInput border) {
    increment(kinds, "border", border != null);
    if (border == null) {
      return;
    }
    increment(kinds, "border_all", border.all() != null);
    increment(kinds, "border_top", border.top() != null);
    increment(kinds, "border_right", border.right() != null);
    increment(kinds, "border_bottom", border.bottom() != null);
    increment(kinds, "border_left", border.left() != null);
    increment(kinds, "border_all_none", isNone(border.all()));
    increment(kinds, "border_top_none", isNone(border.top()));
    increment(kinds, "border_right_none", isNone(border.right()));
    increment(kinds, "border_bottom_none", isNone(border.bottom()));
    increment(kinds, "border_left_none", isNone(border.left()));
    increment(kinds, "border_all_color", hasColor(border.all()));
    increment(kinds, "border_top_color", hasColor(border.top()));
    increment(kinds, "border_right_color", hasColor(border.right()));
    increment(kinds, "border_bottom_color", hasColor(border.bottom()));
    increment(kinds, "border_left_color", hasColor(border.left()));
  }

  private static void appendEngineBorderKinds(Map<String, Long> kinds, ExcelBorder border) {
    increment(kinds, "border", border != null);
    if (border == null) {
      return;
    }
    increment(kinds, "border_all", border.all() != null);
    increment(kinds, "border_top", border.top() != null);
    increment(kinds, "border_right", border.right() != null);
    increment(kinds, "border_bottom", border.bottom() != null);
    increment(kinds, "border_left", border.left() != null);
    increment(kinds, "border_all_none", isNone(border.all()));
    increment(kinds, "border_top_none", isNone(border.top()));
    increment(kinds, "border_right_none", isNone(border.right()));
    increment(kinds, "border_bottom_none", isNone(border.bottom()));
    increment(kinds, "border_left_none", isNone(border.left()));
    increment(kinds, "border_all_color", hasColor(border.all()));
    increment(kinds, "border_top_color", hasColor(border.top()));
    increment(kinds, "border_right_color", hasColor(border.right()));
    increment(kinds, "border_bottom_color", hasColor(border.bottom()));
    increment(kinds, "border_left_color", hasColor(border.left()));
  }

  private static void appendProtocolProtectionKinds(
      Map<String, Long> kinds, CellProtectionInput protection) {
    increment(kinds, "protection", protection != null);
    if (protection == null) {
      return;
    }
    increment(kinds, "locked", protection.locked() != null);
    increment(kinds, "hidden_formula", protection.hiddenFormula() != null);
  }

  private static void appendEngineAlignmentKinds(Map<String, Long> kinds, ExcelCellStyle style) {
    increment(kinds, "alignment", style.alignment() != null);
    if (style.alignment() == null) {
      return;
    }
    increment(kinds, "wrap_text", style.alignment().wrapText() != null);
    increment(kinds, "horizontal_alignment", style.alignment().horizontalAlignment() != null);
    increment(kinds, "vertical_alignment", style.alignment().verticalAlignment() != null);
    increment(kinds, "text_rotation", style.alignment().textRotation() != null);
    increment(kinds, "indentation", style.alignment().indentation() != null);
  }

  private static void appendEngineFontKinds(Map<String, Long> kinds, ExcelCellStyle style) {
    increment(kinds, "font", style.font() != null);
    if (style.font() == null) {
      return;
    }
    increment(kinds, "bold", style.font().bold() != null);
    increment(kinds, "italic", style.font().italic() != null);
    increment(kinds, "font_name", style.font().fontName() != null);
    increment(kinds, "font_height", style.font().fontHeight() != null);
    increment(kinds, "font_color", style.font().fontColor() != null);
    increment(kinds, "underline", style.font().underline() != null);
    increment(kinds, "strikeout", style.font().strikeout() != null);
  }

  private static void appendEngineFillKinds(Map<String, Long> kinds, ExcelCellStyle style) {
    increment(kinds, "fill", style.fill() != null);
    if (style.fill() == null) {
      return;
    }
    increment(kinds, "fill_pattern", style.fill().pattern() != null);
    increment(kinds, "fill_pattern_solid", style.fill().pattern() == ExcelFillPattern.SOLID);
    increment(kinds, "fill_patterned", isPatterned(style.fill().pattern()));
    increment(kinds, "fill_foreground_color", style.fill().foregroundColor() != null);
    increment(kinds, "fill_background_color", style.fill().backgroundColor() != null);
    increment(kinds, "fill_color", style.fill().foregroundColor() != null);
  }

  private static void appendEngineProtectionKinds(Map<String, Long> kinds, ExcelCellStyle style) {
    increment(kinds, "protection", style.protection() != null);
    if (style.protection() == null) {
      return;
    }
    increment(kinds, "locked", style.protection().locked() != null);
    increment(kinds, "hidden_formula", style.protection().hiddenFormula() != null);
  }

  private static boolean isNone(CellBorderSideInput side) {
    return side != null && side.style() == ExcelBorderStyle.NONE;
  }

  private static boolean isNone(ExcelBorderSide side) {
    return side != null && side.style() == dev.erst.gridgrind.excel.ExcelBorderStyle.NONE;
  }

  private static boolean hasColor(CellBorderSideInput side) {
    return side != null && side.color() != null;
  }

  private static boolean hasColor(ExcelBorderSide side) {
    return side != null && side.color() != null;
  }

  private static boolean isPatterned(ExcelFillPattern pattern) {
    return pattern != null && pattern != ExcelFillPattern.NONE && pattern != ExcelFillPattern.SOLID;
  }

  private static void increment(Map<String, Long> counts, String key, boolean present) {
    if (present) {
      counts.merge(key, 1L, Long::sum);
    }
  }
}
