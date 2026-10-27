package dev.erst.gridgrind.fuzz.support;

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
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Extracts stable style-attribute coverage labels from style patches. */
public final class StyleKindIntrospection {
  private StyleKindIntrospection() {}

  /** Returns attribute labels present on a canonical style patch. */
  public static Map<String, Long> styleKinds(CellStylePatch style) {
    Objects.requireNonNull(style, "style must not be null");
    LinkedHashMap<String, Long> kinds = new LinkedHashMap<>();
    increment(kinds, "number_format", style.numberFormat().isPresent());
    appendAlignmentKinds(kinds, style.alignment());
    appendFontKinds(kinds, style.font());
    appendFillKinds(kinds, style.fill());
    appendBorderKinds(kinds, style.border());
    appendProtectionKinds(kinds, style.protection());
    return Map.copyOf(kinds);
  }

  private static void appendAlignmentKinds(
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

  private static void appendFontKinds(Map<String, Long> kinds, Optional<CellFontPatch> font) {
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

  private static void appendFillKinds(Map<String, Long> kinds, Optional<CellFillSpec> fill) {
    increment(kinds, "fill", fill.isPresent());
    if (fill.isEmpty()) {
      return;
    }
    CellFillSpec required = fill.orElseThrow();
    ExcelFillPattern pattern = fillPattern(required);
    increment(kinds, "fill_pattern", pattern != null);
    increment(kinds, "fill_pattern_solid", pattern == ExcelFillPattern.SOLID);
    increment(kinds, "fill_patterned", isPatterned(pattern));
    increment(kinds, "fill_foreground_color", foregroundColor(required) != null);
    increment(kinds, "fill_background_color", backgroundColor(required) != null);
    increment(kinds, "fill_color", foregroundColor(required) != null);
  }

  private static void appendBorderKinds(Map<String, Long> kinds, Optional<CellBorderPatch> border) {
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
    increment(kinds, "border_all_none", isNone(required.all()));
    increment(kinds, "border_top_none", isNone(required.top()));
    increment(kinds, "border_right_none", isNone(required.right()));
    increment(kinds, "border_bottom_none", isNone(required.bottom()));
    increment(kinds, "border_left_none", isNone(required.left()));
    increment(kinds, "border_all_color", hasColor(required.all()));
    increment(kinds, "border_top_color", hasColor(required.top()));
    increment(kinds, "border_right_color", hasColor(required.right()));
    increment(kinds, "border_bottom_color", hasColor(required.bottom()));
    increment(kinds, "border_left_color", hasColor(required.left()));
  }

  private static void appendProtectionKinds(
      Map<String, Long> kinds, Optional<CellProtectionPatch> protection) {
    increment(kinds, "protection", protection.isPresent());
    if (protection.isEmpty()) {
      return;
    }
    CellProtectionPatch required = protection.orElseThrow();
    increment(kinds, "locked", required.locked().isPresent());
    increment(kinds, "hidden_formula", required.hiddenFormula().isPresent());
  }

  private static boolean isNone(Optional<BorderSidePatch> side) {
    return side.flatMap(BorderSidePatch::style)
        .filter(style -> style == ExcelBorderStyle.NONE)
        .isPresent();
  }

  private static boolean hasColor(Optional<BorderSidePatch> side) {
    return side.isPresent() && side.orElseThrow().color().isPresent();
  }

  private static @Nullable ExcelFillPattern fillPattern(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternOnly pattern -> pattern.pattern();
      case CellFillSpec.PatternForeground pattern -> pattern.pattern();
      case CellFillSpec.PatternBackground pattern -> pattern.pattern();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.pattern();
      case CellFillSpec.Gradient ignored -> null;
    };
  }

  private static @Nullable Object foregroundColor(CellFillSpec fill) {
    return switch (fill) {
      case CellFillSpec.PatternForeground pattern -> pattern.foregroundColor();
      case CellFillSpec.PatternForegroundBackground pattern -> pattern.foregroundColor();
      default -> null;
    };
  }

  private static @Nullable Object backgroundColor(CellFillSpec fill) {
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
