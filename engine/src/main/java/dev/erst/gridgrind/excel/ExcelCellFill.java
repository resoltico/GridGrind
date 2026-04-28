package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import java.util.Objects;

/** Fill patch applied through {@link ExcelCellStyle}. */
public sealed interface ExcelCellFill
    permits ExcelCellFill.PatternOnly,
        ExcelCellFill.PatternForeground,
        ExcelCellFill.PatternBackground,
        ExcelCellFill.PatternForegroundBackground,
        ExcelCellFill.Gradient {
  /** Creates one pattern-only fill patch. */
  static PatternOnly pattern(ExcelFillPattern pattern) {
    return new PatternOnly(pattern);
  }

  /** Creates one patterned fill patch carrying foreground color semantics. */
  static PatternForeground patternForeground(ExcelFillPattern pattern, ExcelColor foregroundColor) {
    return new PatternForeground(pattern, foregroundColor);
  }

  /** Creates one patterned fill patch carrying background color semantics. */
  static PatternBackground patternBackground(ExcelFillPattern pattern, ExcelColor backgroundColor) {
    return new PatternBackground(pattern, backgroundColor);
  }

  /** Creates one patterned fill patch carrying both foreground and background colors. */
  static PatternForegroundBackground patternColors(
      ExcelFillPattern pattern, ExcelColor foregroundColor, ExcelColor backgroundColor) {
    return new PatternForegroundBackground(pattern, foregroundColor, backgroundColor);
  }

  /** Creates one gradient fill patch. */
  static Gradient gradient(ExcelGradientFill gradient) {
    return new Gradient(gradient);
  }

  /** Pattern-only fill patch with no explicit colors. */
  record PatternOnly(ExcelFillPattern pattern) implements ExcelCellFill {
    public PatternOnly {
      Objects.requireNonNull(pattern, "pattern must not be null");
    }
  }

  /** Patterned fill patch carrying one foreground color. */
  record PatternForeground(ExcelFillPattern pattern, ExcelColor foregroundColor)
      implements ExcelCellFill {
    public PatternForeground {
      requireForegroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
    }
  }

  /** Patterned fill patch carrying one background color. */
  record PatternBackground(ExcelFillPattern pattern, ExcelColor backgroundColor)
      implements ExcelCellFill {
    public PatternBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Patterned fill patch carrying foreground and background colors. */
  record PatternForegroundBackground(
      ExcelFillPattern pattern, ExcelColor foregroundColor, ExcelColor backgroundColor)
      implements ExcelCellFill {
    public PatternForegroundBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Gradient fill patch. */
  record Gradient(ExcelGradientFill gradient) implements ExcelCellFill {
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
