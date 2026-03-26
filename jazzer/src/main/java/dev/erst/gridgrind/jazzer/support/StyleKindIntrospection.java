package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.ExcelBorder;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.protocol.CellBorderInput;
import dev.erst.gridgrind.protocol.CellBorderSideInput;
import dev.erst.gridgrind.protocol.CellStyleInput;
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
    increment(kinds, "bold", style.bold() != null);
    increment(kinds, "italic", style.italic() != null);
    increment(kinds, "wrap_text", style.wrapText() != null);
    increment(kinds, "horizontal_alignment", style.horizontalAlignment() != null);
    increment(kinds, "vertical_alignment", style.verticalAlignment() != null);
    increment(kinds, "font_name", style.fontName() != null);
    increment(kinds, "font_height", style.fontHeight() != null);
    increment(kinds, "font_height_points", style.fontHeight() instanceof dev.erst.gridgrind.protocol.FontHeightInput.Points);
    increment(kinds, "font_height_twips", style.fontHeight() instanceof dev.erst.gridgrind.protocol.FontHeightInput.Twips);
    increment(kinds, "font_color", style.fontColor() != null);
    increment(kinds, "underline", style.underline() != null);
    increment(kinds, "strikeout", style.strikeout() != null);
    increment(kinds, "fill_color", style.fillColor() != null);
    appendProtocolBorderKinds(kinds, style.border());
    return Map.copyOf(kinds);
  }

  /** Returns attribute labels present on an engine style patch. */
  public static Map<String, Long> styleKinds(ExcelCellStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    LinkedHashMap<String, Long> kinds = new LinkedHashMap<>();
    increment(kinds, "number_format", style.numberFormat() != null);
    increment(kinds, "bold", style.bold() != null);
    increment(kinds, "italic", style.italic() != null);
    increment(kinds, "wrap_text", style.wrapText() != null);
    increment(kinds, "horizontal_alignment", style.horizontalAlignment() != null);
    increment(kinds, "vertical_alignment", style.verticalAlignment() != null);
    increment(kinds, "font_name", style.fontName() != null);
    increment(kinds, "font_height", style.fontHeight() != null);
    increment(kinds, "font_color", style.fontColor() != null);
    increment(kinds, "underline", style.underline() != null);
    increment(kinds, "strikeout", style.strikeout() != null);
    increment(kinds, "fill_color", style.fillColor() != null);
    appendEngineBorderKinds(kinds, style.border());
    return Map.copyOf(kinds);
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
  }

  private static boolean isNone(CellBorderSideInput side) {
    return side != null && side.style() == dev.erst.gridgrind.excel.ExcelBorderStyle.NONE;
  }

  private static boolean isNone(ExcelBorderSide side) {
    return side != null && side.style() == dev.erst.gridgrind.excel.ExcelBorderStyle.NONE;
  }

  private static void increment(Map<String, Long> counts, String key, boolean present) {
    if (present) {
      counts.merge(key, 1L, Long::sum);
    }
  }
}
