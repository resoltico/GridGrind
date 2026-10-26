package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellFontPatch;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelCellFont;
import dev.erst.gridgrind.excel.ExcelCellStyle;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Extracts stable style-attribute coverage labels from style patches. */
public final class StyleKindIntrospection {
  private StyleKindIntrospection() {}

  /** Returns attribute labels present on a protocol style patch. */
  public static Map<String, Long> styleKinds(CellStylePatch style) {
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
      Map<String, Long> kinds, Optional<CellAlignmentPatch> alignment) {
    increment(kinds, "alignment", alignment.isPresent());
    if (alignment.isEmpty()) {
      return;
    }
    CellAlignmentPatch required = alignment.orElseThrow();
    increment(kinds, "wrap_text", required.wrapText().isPresent());
    increment(kinds, "horizontal_alignment", required.horizontalAlignment().isPresent());
    increment(kinds, "vertical_alignment", required.verticalAlignment().isPresent());
    increment(kinds, "text_rotation", required.textRotation().isPresent());
    increment(kinds, "indentation", required.indentation().isPresent());
  }

  private static void appendProtocolFontKinds(
      Map<String, Long> kinds, Optional<CellFontPatch> font) {
    increment(kinds, "font", font.isPresent());
    if (font.isEmpty()) {
      return;
    }
    CellFontPatch required = font.orElseThrow();
    increment(kinds, "bold", required.bold().isPresent());
    increment(kinds, "italic", required.italic().isPresent());
    increment(kinds, "font_name", required.fontName().isPresent());
    increment(kinds, "font_height", required.fontHeight().isPresent());
    increment(
        kinds,
        "font_height_points",
        required.fontHeight().filter(FontHeightSpec.Points.class::isInstance).isPresent());
    increment(
        kinds,
        "font_height_twips",
        required.fontHeight().filter(FontHeightSpec.Twips.class::isInstance).isPresent());
    increment(kinds, "font_color", required.fontColor().isPresent());
    increment(kinds, "underline", required.underline().isPresent());
    increment(kinds, "strikeout", required.strikeout().isPresent());
  }

  private static void appendProtocolFillKinds(
      Map<String, Long> kinds, Optional<CellFillSpec> fill) {
    increment(kinds, "fill", fill.isPresent());
    if (fill.isEmpty()) {
      return;
    }
    CellFillSpec required = fill.orElseThrow();
    ExcelFillPattern pattern = protocolFillPattern(required);
    increment(kinds, "fill_pattern", pattern != null);
    increment(kinds, "fill_pattern_solid", pattern == ExcelFillPattern.SOLID);
    increment(kinds, "fill_patterned", isPatterned(pattern));
    increment(kinds, "fill_foreground_color", protocolForegroundColor(required) != null);
    increment(kinds, "fill_background_color", protocolBackgroundColor(required) != null);
    increment(kinds, "fill_color", protocolForegroundColor(required) != null);
  }

  private static void appendProtocolBorderKinds(
      Map<String, Long> kinds, Optional<CellBorderPatch> border) {
    increment(kinds, "border", border.isPresent());
    if (border.isEmpty()) {
      return;
    }
    CellBorderPatch required = border.orElseThrow();
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
      Map<String, Long> kinds, Optional<CellBorderPatch> border) {
    increment(kinds, "border", border.isPresent());
    if (border.isEmpty()) {
      return;
    }
    CellBorderPatch required = border.orElseThrow();
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
      Map<String, Long> kinds, Optional<CellProtectionPatch> protection) {
    increment(kinds, "protection", protection.isPresent());
    if (protection.isEmpty()) {
      return;
    }
    CellProtectionPatch required = protection.orElseThrow();
    increment(kinds, "locked", required.locked().isPresent());
    increment(kinds, "hidden_formula", required.hiddenFormula().isPresent());
  }

  private static void appendEngineAlignmentKinds(Map<String, Long> kinds, ExcelCellStyle style) {
    increment(kinds, "alignment", style.alignment().isPresent());
    if (style.alignment().isEmpty()) {
      return;
    }
    CellAlignmentPatch alignment = style.alignment().orElseThrow();
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
    CellFillSpec fill = style.fill().orElseThrow();
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
    CellProtectionPatch protection = style.protection().orElseThrow();
    increment(kinds, "locked", protection.locked().isPresent());
    increment(kinds, "hidden_formula", protection.hiddenFormula().isPresent());
  }

  private static boolean isProtocolNone(Optional<BorderSidePatch> side) {
    return side.isPresent() && side.orElseThrow().style().orElseThrow() == ExcelBorderStyle.NONE;
  }

  private static boolean isEngineNone(Optional<BorderSidePatch> side) {
    return side.isPresent()
        && side.orElseThrow().style().orElseThrow()
            == dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle.NONE;
  }

  private static boolean hasProtocolColor(Optional<BorderSidePatch> side) {
    return side.isPresent() && side.orElseThrow().color().isPresent();
  }

  private static boolean hasEngineColor(Optional<BorderSidePatch> side) {
    return side.isPresent() && side.orElseThrow().color().isPresent();
  }

  private static @Nullable ExcelFillPattern protocolFillPattern(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternOnly pattern -> pattern.pattern();
      case CellFillSpec.PatternForeground pattern -> pattern.pattern();
      case CellFillSpec.PatternBackground pattern -> pattern.pattern();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.pattern();
      case CellFillSpec.Gradient ignored -> null;
    };
  }

  private static @Nullable Object protocolForegroundColor(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternForeground pattern -> pattern.foregroundColor();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.foregroundColor();
      default -> null;
    };
  }

  private static @Nullable Object protocolBackgroundColor(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternBackground pattern -> pattern.backgroundColor();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.backgroundColor();
      default -> null;
    };
  }

  private static @Nullable ExcelFillPattern engineFillPattern(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternOnly pattern -> pattern.pattern();
      case CellFillSpec.PatternForeground pattern -> pattern.pattern();
      case CellFillSpec.PatternBackground pattern -> pattern.pattern();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.pattern();
      case CellFillSpec.Gradient ignored -> null;
    };
  }

  private static @Nullable Object engineForegroundColor(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternForeground pattern -> pattern.foregroundColor();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.foregroundColor();
      default -> null;
    };
  }

  private static @Nullable Object engineBackgroundColor(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternBackground pattern -> pattern.backgroundColor();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.backgroundColor();
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
