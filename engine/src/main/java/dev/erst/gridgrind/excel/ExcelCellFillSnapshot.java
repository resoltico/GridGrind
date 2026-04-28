package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import java.util.Objects;

/** Immutable snapshot of the fill currently applied to a cell. */
public sealed interface ExcelCellFillSnapshot
    permits ExcelCellFillSnapshot.PatternOnly,
        ExcelCellFillSnapshot.PatternForeground,
        ExcelCellFillSnapshot.PatternBackground,
        ExcelCellFillSnapshot.PatternForegroundBackground,
        ExcelCellFillSnapshot.Gradient {
  /** Returns one pattern-only fill snapshot. */
  static PatternOnly pattern(ExcelFillPattern pattern) {
    return new PatternOnly(pattern);
  }

  /** Returns one patterned fill snapshot carrying foreground color metadata. */
  static PatternForeground patternForeground(
      ExcelFillPattern pattern, ExcelColorSnapshot foregroundColor) {
    return new PatternForeground(pattern, foregroundColor);
  }

  /** Returns one patterned fill snapshot carrying background color metadata. */
  static PatternBackground patternBackground(
      ExcelFillPattern pattern, ExcelColorSnapshot backgroundColor) {
    return new PatternBackground(pattern, backgroundColor);
  }

  /** Returns one patterned fill snapshot carrying foreground and background colors. */
  static PatternForegroundBackground patternColors(
      ExcelFillPattern pattern,
      ExcelColorSnapshot foregroundColor,
      ExcelColorSnapshot backgroundColor) {
    return new PatternForegroundBackground(pattern, foregroundColor, backgroundColor);
  }

  /** Returns one gradient fill snapshot. */
  static Gradient gradient(ExcelGradientFillSnapshot gradient) {
    return new Gradient(gradient);
  }

  /** Pattern-only fill snapshot. */
  record PatternOnly(ExcelFillPattern pattern) implements ExcelCellFillSnapshot {
    public PatternOnly {
      Objects.requireNonNull(pattern, "pattern must not be null");
    }
  }

  /** Patterned fill snapshot carrying one foreground color. */
  record PatternForeground(ExcelFillPattern pattern, ExcelColorSnapshot foregroundColor)
      implements ExcelCellFillSnapshot {
    public PatternForeground {
      requireForegroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
    }
  }

  /** Patterned fill snapshot carrying one background color. */
  record PatternBackground(ExcelFillPattern pattern, ExcelColorSnapshot backgroundColor)
      implements ExcelCellFillSnapshot {
    public PatternBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Patterned fill snapshot carrying foreground and background colors. */
  record PatternForegroundBackground(
      ExcelFillPattern pattern,
      ExcelColorSnapshot foregroundColor,
      ExcelColorSnapshot backgroundColor)
      implements ExcelCellFillSnapshot {
    public PatternForegroundBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Gradient fill snapshot. */
  record Gradient(ExcelGradientFillSnapshot gradient) implements ExcelCellFillSnapshot {
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
