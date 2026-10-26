package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import java.util.Objects;

/** Protocol-facing fill patch used by {@link CellStylePatch}. */
@Union(discriminator = "type", catalog = true, displayOrder = 57)
public sealed interface CellFillSpec
    permits dev.erst.gridgrind.core.model.style.CellFillSpec.PatternOnly,
        dev.erst.gridgrind.core.model.style.CellFillSpec.PatternForeground,
        dev.erst.gridgrind.core.model.style.CellFillSpec.PatternBackground,
        dev.erst.gridgrind.core.model.style.CellFillSpec.PatternForegroundBackground,
        dev.erst.gridgrind.core.model.style.CellFillSpec.Gradient {
  /** Creates one pattern-only fill patch. */
  static PatternOnly pattern(ExcelFillPattern pattern) {
    return new PatternOnly(pattern);
  }

  /** Creates one patterned fill patch with foreground color semantics. */
  static PatternForeground patternForeground(ExcelFillPattern pattern, ColorSpec foregroundColor) {
    return new PatternForeground(pattern, foregroundColor);
  }

  /** Creates one patterned fill patch with background color semantics. */
  static PatternBackground patternBackground(ExcelFillPattern pattern, ColorSpec backgroundColor) {
    return new PatternBackground(pattern, backgroundColor);
  }

  /** Creates one patterned fill patch with both foreground and background colors. */
  static PatternForegroundBackground patternColors(
      ExcelFillPattern pattern, ColorSpec foregroundColor, ColorSpec backgroundColor) {
    return new PatternForegroundBackground(pattern, foregroundColor, backgroundColor);
  }

  /** Creates one gradient fill patch. */
  static Gradient gradient(CellGradientFillSpec gradient) {
    return new Gradient(gradient);
  }

  /** Pattern-only fill patch with no explicit colors. */
  @Variant(
      id = "PATTERN_ONLY",
      summary = "Write one patterned fill with no explicit foreground or background colors.")
  record PatternOnly(ExcelFillPattern pattern) implements CellFillSpec {
    public PatternOnly {
      Objects.requireNonNull(pattern, "pattern must not be null");
    }
  }

  /** Patterned fill patch carrying one foreground color. */
  @Variant(
      id = "PATTERN_FOREGROUND",
      summary = "Write one patterned fill with one explicit foreground color.")
  record PatternForeground(ExcelFillPattern pattern, ColorSpec foregroundColor)
      implements CellFillSpec {
    public PatternForeground {
      requireForegroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
    }
  }

  /** Patterned fill patch carrying one background color. */
  @Variant(
      id = "PATTERN_BACKGROUND",
      summary = "Write one patterned fill with one explicit background color.")
  record PatternBackground(ExcelFillPattern pattern, ColorSpec backgroundColor)
      implements CellFillSpec {
    public PatternBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Patterned fill patch carrying both foreground and background colors. */
  @Variant(
      id = "PATTERN_FOREGROUND_BACKGROUND",
      summary = "Write one patterned fill with explicit foreground and background colors.")
  record PatternForegroundBackground(
      ExcelFillPattern pattern, ColorSpec foregroundColor, ColorSpec backgroundColor)
      implements CellFillSpec {
    public PatternForegroundBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Gradient fill patch. */
  @Variant(id = "GRADIENT", summary = "Write one gradient fill patch.")
  record Gradient(CellGradientFillSpec gradient) implements CellFillSpec {
    public Gradient {
      Objects.requireNonNull(gradient, "gradient must not be null");
    }
  }

  private static void requireForegroundPattern(ExcelFillPattern pattern) {
    Objects.requireNonNull(pattern, "pattern must not be null");
    if (pattern == ExcelFillPattern.NONE) {
      throw new IllegalArgumentException("fill pattern NONE does not accept colors");
    }
  }

  private static void requireBackgroundPattern(ExcelFillPattern pattern) {
    requireForegroundPattern(pattern);
    if (pattern == ExcelFillPattern.SOLID) {
      throw new IllegalArgumentException("fill backgroundColor is not supported for SOLID fills");
    }
  }
}
