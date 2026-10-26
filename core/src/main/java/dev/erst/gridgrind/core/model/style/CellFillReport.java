package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import java.util.Objects;

/** Effective fill facts reported with every analyzed cell. */
@Union(discriminator = "kind", catalog = true, displayOrder = 61)
public sealed interface CellFillReport
    permits dev.erst.gridgrind.core.model.style.CellFillReport.PatternOnly,
        dev.erst.gridgrind.core.model.style.CellFillReport.PatternForeground,
        dev.erst.gridgrind.core.model.style.CellFillReport.PatternBackground,
        dev.erst.gridgrind.core.model.style.CellFillReport.PatternForegroundBackground,
        dev.erst.gridgrind.core.model.style.CellFillReport.Gradient {
  /** Returns one pattern-only fill report. */
  static PatternOnly pattern(ExcelFillPattern pattern) {
    return new PatternOnly(pattern);
  }

  /** Returns one patterned fill report carrying foreground color metadata. */
  static PatternForeground patternForeground(
      ExcelFillPattern pattern, CellColorReport foregroundColor) {
    return new PatternForeground(pattern, foregroundColor);
  }

  /** Returns one patterned fill report carrying background color metadata. */
  static PatternBackground patternBackground(
      ExcelFillPattern pattern, CellColorReport backgroundColor) {
    return new PatternBackground(pattern, backgroundColor);
  }

  /** Returns one patterned fill report carrying foreground and background colors. */
  static PatternForegroundBackground patternColors(
      ExcelFillPattern pattern, CellColorReport foregroundColor, CellColorReport backgroundColor) {
    return new PatternForegroundBackground(pattern, foregroundColor, backgroundColor);
  }

  /** Returns one gradient fill report. */
  static Gradient gradient(CellGradientFillReport gradient) {
    return new Gradient(gradient);
  }

  /** Pattern-only fill report. */
  @Variant(
      id = "PATTERN_ONLY",
      summary = "Read one patterned fill with no explicit foreground or background colors.")
  record PatternOnly(ExcelFillPattern pattern) implements CellFillReport {
    public PatternOnly {
      Objects.requireNonNull(pattern, "pattern must not be null");
    }
  }

  /** Patterned fill report carrying one foreground color. */
  @Variant(
      id = "PATTERN_FOREGROUND",
      summary = "Read one patterned fill with one factual foreground color.")
  record PatternForeground(ExcelFillPattern pattern, CellColorReport foregroundColor)
      implements CellFillReport {
    public PatternForeground {
      requireForegroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
    }
  }

  /** Patterned fill report carrying one background color. */
  @Variant(
      id = "PATTERN_BACKGROUND",
      summary = "Read one patterned fill with one factual background color.")
  record PatternBackground(ExcelFillPattern pattern, CellColorReport backgroundColor)
      implements CellFillReport {
    public PatternBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Patterned fill report carrying foreground and background colors. */
  @Variant(
      id = "PATTERN_FOREGROUND_BACKGROUND",
      summary = "Read one patterned fill with factual foreground and background colors.")
  record PatternForegroundBackground(
      ExcelFillPattern pattern, CellColorReport foregroundColor, CellColorReport backgroundColor)
      implements CellFillReport {
    public PatternForegroundBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Gradient fill report. */
  @Variant(id = "GRADIENT", summary = "Read one factual gradient fill.")
  record Gradient(CellGradientFillReport gradient) implements CellFillReport {
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
