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
import dev.erst.gridgrind.excel.ExcelCellAlignment;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellProtection;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Extracts stable style-attribute coverage labels from style patches. */
public final class StyleKindIntrospection {
  private StyleKindIntrospection() {}

  /** Returns attribute labels present on a protocol style patch. */
  public static Map<String, Long> styleKinds(CellStyleInput style) {
    Objects.requireNonNull(style, "style must not be null");
    LinkedHashMap<String, Long> kinds = new LinkedHashMap<>();
    increment(kinds, "number_format", style.numberFormat().isPresent());
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
    increment(kinds, "number_format", style.numberFormat().isPresent());
    appendEngineAlignmentKinds(kinds, style);
    appendEngineFontKinds(kinds, style);
    appendEngineFillKinds(kinds, style);
    appendEngineBorderKinds(kinds, style.border());
    appendEngineProtectionKinds(kinds, style);
    return Map.copyOf(kinds);
  }

  private static void appendProtocolAlignmentKinds(
      Map<String, Long> kinds, Optional<CellAlignmentInput> alignment) {
    increment(kinds, "alignment", alignment.isPresent());
    if (alignment.isEmpty()) {
      return;
    }
    CellAlignmentInput required = alignment.orElseThrow();
    increment(kinds, "wrap_text", required.wrapText().isPresent());
    increment(kinds, "horizontal_alignment", required.horizontalAlignment().isPresent());
    increment(kinds, "vertical_alignment", required.verticalAlignment().isPresent());
    increment(kinds, "text_rotation", required.textRotation().isPresent());
    increment(kinds, "indentation", required.indentation().isPresent());
  }

  private static void appendProtocolFontKinds(
      Map<String, Long> kinds, Optional<CellFontInput> font) {
    increment(kinds, "font", font.isPresent());
    if (font.isEmpty()) {
      return;
    }
    CellFontInput required = font.orElseThrow();
    increment(kinds, "bold", required.bold().isPresent());
    increment(kinds, "italic", required.italic().isPresent());
    increment(kinds, "font_name", required.fontName().isPresent());
    increment(kinds, "font_height", required.fontHeight().isPresent());
    increment(
        kinds,
        "font_height_points",
        required.fontHeight().filter(FontHeightInput.Points.class::isInstance).isPresent());
    increment(
        kinds,
        "font_height_twips",
        required.fontHeight().filter(FontHeightInput.Twips.class::isInstance).isPresent());
    increment(kinds, "font_color", required.fontColor().isPresent());
    increment(kinds, "underline", required.underline().isPresent());
    increment(kinds, "strikeout", required.strikeout().isPresent());
  }

  private static void appendProtocolFillKinds(
      Map<String, Long> kinds, Optional<CellFillInput> fill) {
    increment(kinds, "fill", fill.isPresent());
    if (fill.isEmpty()) {
      return;
    }
    CellFillInput required = fill.orElseThrow();
    ExcelFillPattern pattern = protocolFillPattern(required);
    increment(kinds, "fill_pattern", pattern != null);
    increment(kinds, "fill_pattern_solid", pattern == ExcelFillPattern.SOLID);
    increment(kinds, "fill_patterned", isPatterned(pattern));
    increment(kinds, "fill_foreground_color", protocolForegroundColor(required) != null);
    increment(kinds, "fill_background_color", protocolBackgroundColor(required) != null);
    increment(kinds, "fill_color", protocolForegroundColor(required) != null);
  }

  private static void appendProtocolBorderKinds(
      Map<String, Long> kinds, Optional<CellBorderInput> border) {
    increment(kinds, "border", border.isPresent());
    if (border.isEmpty()) {
      return;
    }
    CellBorderInput required = border.orElseThrow();
    increment(kinds, "border_all", required.all().isPresent());
    increment(kinds, "border_top", required.top().isPresent());
    increment(kinds, "border_right", required.right().isPresent());
    increment(kinds, "border_bottom", required.bottom().isPresent());
    increment(kinds, "border_left", required.left().isPresent());
    increment(kinds, "border_all_none", isProtocolNone(required.all()));
    increment(kinds, "border_top_none", isProtocolNone(required.top()));
    increment(kinds, "border_right_none", isProtocolNone(required.right()));
    increment(kinds, "border_bottom_none", isProtocolNone(required.bottom()));
    increment(kinds, "border_left_none", isProtocolNone(required.left()));
    increment(kinds, "border_all_color", hasProtocolColor(required.all()));
    increment(kinds, "border_top_color", hasProtocolColor(required.top()));
    increment(kinds, "border_right_color", hasProtocolColor(required.right()));
    increment(kinds, "border_bottom_color", hasProtocolColor(required.bottom()));
    increment(kinds, "border_left_color", hasProtocolColor(required.left()));
  }

  private static void appendEngineBorderKinds(
      Map<String, Long> kinds, Optional<ExcelBorder> border) {
    increment(kinds, "border", border.isPresent());
    if (border.isEmpty()) {
      return;
    }
    ExcelBorder required = border.orElseThrow();
    increment(kinds, "border_all", required.all().isPresent());
    increment(kinds, "border_top", required.top().isPresent());
    increment(kinds, "border_right", required.right().isPresent());
    increment(kinds, "border_bottom", required.bottom().isPresent());
    increment(kinds, "border_left", required.left().isPresent());
    increment(kinds, "border_all_none", isEngineNone(required.all()));
    increment(kinds, "border_top_none", isEngineNone(required.top()));
    increment(kinds, "border_right_none", isEngineNone(required.right()));
    increment(kinds, "border_bottom_none", isEngineNone(required.bottom()));
    increment(kinds, "border_left_none", isEngineNone(required.left()));
    increment(kinds, "border_all_color", hasEngineColor(required.all()));
    increment(kinds, "border_top_color", hasEngineColor(required.top()));
    increment(kinds, "border_right_color", hasEngineColor(required.right()));
    increment(kinds, "border_bottom_color", hasEngineColor(required.bottom()));
    increment(kinds, "border_left_color", hasEngineColor(required.left()));
  }

  private static void appendProtocolProtectionKinds(
      Map<String, Long> kinds, Optional<CellProtectionInput> protection) {
    increment(kinds, "protection", protection.isPresent());
    if (protection.isEmpty()) {
      return;
    }
    CellProtectionInput required = protection.orElseThrow();
    increment(kinds, "locked", required.locked().isPresent());
    increment(kinds, "hidden_formula", required.hiddenFormula().isPresent());
  }

  private static void appendEngineAlignmentKinds(Map<String, Long> kinds, ExcelCellStyle style) {
    increment(kinds, "alignment", style.alignment().isPresent());
    if (style.alignment().isEmpty()) {
      return;
    }
    ExcelCellAlignment alignment = style.alignment().orElseThrow();
    increment(kinds, "wrap_text", alignment.wrapText().isPresent());
    increment(kinds, "horizontal_alignment", alignment.horizontalAlignment().isPresent());
    increment(kinds, "vertical_alignment", alignment.verticalAlignment().isPresent());
    increment(kinds, "text_rotation", alignment.textRotation().isPresent());
    increment(kinds, "indentation", alignment.indentation().isPresent());
  }

  private static void appendEngineFontKinds(Map<String, Long> kinds, ExcelCellStyle style) {
    increment(kinds, "font", style.font().isPresent());
    if (style.font().isEmpty()) {
      return;
    }
    ExcelCellFont font = style.font().orElseThrow();
    increment(kinds, "bold", font.bold().isPresent());
    increment(kinds, "italic", font.italic().isPresent());
    increment(kinds, "font_name", font.fontName().isPresent());
    increment(kinds, "font_height", font.fontHeight().isPresent());
    increment(kinds, "font_color", font.fontColor().isPresent());
    increment(kinds, "underline", font.underline().isPresent());
    increment(kinds, "strikeout", font.strikeout().isPresent());
  }

  private static void appendEngineFillKinds(Map<String, Long> kinds, ExcelCellStyle style) {
    increment(kinds, "fill", style.fill().isPresent());
    if (style.fill().isEmpty()) {
      return;
    }
    ExcelCellFill fill = style.fill().orElseThrow();
    ExcelFillPattern pattern = engineFillPattern(fill);
    increment(kinds, "fill_pattern", pattern != null);
    increment(kinds, "fill_pattern_solid", pattern == ExcelFillPattern.SOLID);
    increment(kinds, "fill_patterned", isPatterned(pattern));
    increment(kinds, "fill_foreground_color", engineForegroundColor(fill) != null);
    increment(kinds, "fill_background_color", engineBackgroundColor(fill) != null);
    increment(kinds, "fill_color", engineForegroundColor(fill) != null);
  }

  private static void appendEngineProtectionKinds(Map<String, Long> kinds, ExcelCellStyle style) {
    increment(kinds, "protection", style.protection().isPresent());
    if (style.protection().isEmpty()) {
      return;
    }
    ExcelCellProtection protection = style.protection().orElseThrow();
    increment(kinds, "locked", protection.locked().isPresent());
    increment(kinds, "hidden_formula", protection.hiddenFormula().isPresent());
  }

  private static boolean isProtocolNone(Optional<CellBorderSideInput> side) {
    return side.isPresent() && side.orElseThrow().style().orElseThrow() == ExcelBorderStyle.NONE;
  }

  private static boolean isEngineNone(Optional<ExcelBorderSide> side) {
    return side.isPresent()
        && side.orElseThrow().style().orElseThrow()
            == dev.erst.gridgrind.excel.foundation.ExcelBorderStyle.NONE;
  }

  private static boolean hasProtocolColor(Optional<CellBorderSideInput> side) {
    return side.isPresent() && side.orElseThrow().color().isPresent();
  }

  private static boolean hasEngineColor(Optional<ExcelBorderSide> side) {
    return side.isPresent() && side.orElseThrow().color().isPresent();
  }

  private static @Nullable ExcelFillPattern protocolFillPattern(CellFillInput fill) {
    return switch (fill) {
      case CellFillInput.PatternOnly pattern -> pattern.pattern();
      case CellFillInput.PatternForeground pattern -> pattern.pattern();
      case CellFillInput.PatternBackground pattern -> pattern.pattern();
      case CellFillInput.PatternForegroundBackground pattern -> pattern.pattern();
      case CellFillInput.Gradient ignored -> null;
    };
  }

  private static @Nullable Object protocolForegroundColor(CellFillInput fill) {
    return switch (fill) {
      case CellFillInput.PatternForeground pattern -> pattern.foregroundColor();
      case CellFillInput.PatternForegroundBackground pattern -> pattern.foregroundColor();
      default -> null;
    };
  }

  private static @Nullable Object protocolBackgroundColor(CellFillInput fill) {
    return switch (fill) {
      case CellFillInput.PatternBackground pattern -> pattern.backgroundColor();
      case CellFillInput.PatternForegroundBackground pattern -> pattern.backgroundColor();
      default -> null;
    };
  }

  private static @Nullable ExcelFillPattern engineFillPattern(ExcelCellFill fill) {
    return switch (fill) {
      case ExcelCellFill.PatternOnly pattern -> pattern.pattern();
      case ExcelCellFill.PatternForeground pattern -> pattern.pattern();
      case ExcelCellFill.PatternBackground pattern -> pattern.pattern();
      case ExcelCellFill.PatternForegroundBackground pattern -> pattern.pattern();
      case ExcelCellFill.Gradient ignored -> null;
    };
  }

  private static @Nullable Object engineForegroundColor(ExcelCellFill fill) {
    return switch (fill) {
      case ExcelCellFill.PatternForeground pattern -> pattern.foregroundColor();
      case ExcelCellFill.PatternForegroundBackground pattern -> pattern.foregroundColor();
      default -> null;
    };
  }

  private static @Nullable Object engineBackgroundColor(ExcelCellFill fill) {
    return switch (fill) {
      case ExcelCellFill.PatternBackground pattern -> pattern.backgroundColor();
      case ExcelCellFill.PatternForegroundBackground pattern -> pattern.backgroundColor();
      default -> null;
    };
  }

  private static boolean isPatterned(@Nullable ExcelFillPattern pattern) {
    return pattern != null && pattern != ExcelFillPattern.NONE && pattern != ExcelFillPattern.SOLID;
  }

  private static void increment(Map<String, Long> counts, String key, boolean present) {
    if (present) {
      counts.merge(key, 1L, Long::sum);
    }
  }
}
