package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import java.util.Objects;

/** Effective fill facts reported with every analyzed cell. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellFillReport.PatternOnly.class, name = "PATTERN_ONLY"),
  @JsonSubTypes.Type(value = CellFillReport.PatternForeground.class, name = "PATTERN_FOREGROUND"),
  @JsonSubTypes.Type(value = CellFillReport.PatternBackground.class, name = "PATTERN_BACKGROUND"),
  @JsonSubTypes.Type(
      value = CellFillReport.PatternForegroundBackground.class,
      name = "PATTERN_FOREGROUND_BACKGROUND"),
  @JsonSubTypes.Type(value = CellFillReport.Gradient.class, name = "GRADIENT")
})
public sealed interface CellFillReport
    permits CellFillReport.PatternOnly,
        CellFillReport.PatternForeground,
        CellFillReport.PatternBackground,
        CellFillReport.PatternForegroundBackground,
        CellFillReport.Gradient {
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
  record PatternOnly(ExcelFillPattern pattern) implements CellFillReport {
    public PatternOnly {
      Objects.requireNonNull(pattern, "pattern must not be null");
    }
  }

  /** Patterned fill report carrying one foreground color. */
  record PatternForeground(ExcelFillPattern pattern, CellColorReport foregroundColor)
      implements CellFillReport {
    public PatternForeground {
      requireForegroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
    }
  }

  /** Patterned fill report carrying one background color. */
  record PatternBackground(ExcelFillPattern pattern, CellColorReport backgroundColor)
      implements CellFillReport {
    public PatternBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Patterned fill report carrying foreground and background colors. */
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
