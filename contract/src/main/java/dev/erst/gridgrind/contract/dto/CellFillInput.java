package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import java.util.Objects;

/** Protocol-facing fill patch used by {@link CellStyleInput}. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellFillInput.PatternOnly.class, name = "PATTERN_ONLY"),
  @JsonSubTypes.Type(value = CellFillInput.PatternForeground.class, name = "PATTERN_FOREGROUND"),
  @JsonSubTypes.Type(value = CellFillInput.PatternBackground.class, name = "PATTERN_BACKGROUND"),
  @JsonSubTypes.Type(
      value = CellFillInput.PatternForegroundBackground.class,
      name = "PATTERN_FOREGROUND_BACKGROUND"),
  @JsonSubTypes.Type(value = CellFillInput.Gradient.class, name = "GRADIENT")
})
public sealed interface CellFillInput
    permits CellFillInput.PatternOnly,
        CellFillInput.PatternForeground,
        CellFillInput.PatternBackground,
        CellFillInput.PatternForegroundBackground,
        CellFillInput.Gradient {
  /** Creates one pattern-only fill patch. */
  static PatternOnly pattern(ExcelFillPattern pattern) {
    return new PatternOnly(pattern);
  }

  /** Creates one patterned fill patch with foreground color semantics. */
  static PatternForeground patternForeground(ExcelFillPattern pattern, ColorInput foregroundColor) {
    return new PatternForeground(pattern, foregroundColor);
  }

  /** Creates one patterned fill patch with background color semantics. */
  static PatternBackground patternBackground(ExcelFillPattern pattern, ColorInput backgroundColor) {
    return new PatternBackground(pattern, backgroundColor);
  }

  /** Creates one patterned fill patch with both foreground and background colors. */
  static PatternForegroundBackground patternColors(
      ExcelFillPattern pattern, ColorInput foregroundColor, ColorInput backgroundColor) {
    return new PatternForegroundBackground(pattern, foregroundColor, backgroundColor);
  }

  /** Creates one gradient fill patch. */
  static Gradient gradient(CellGradientFillInput gradient) {
    return new Gradient(gradient);
  }

  /** Pattern-only fill patch with no explicit colors. */
  record PatternOnly(ExcelFillPattern pattern) implements CellFillInput {
    public PatternOnly {
      Objects.requireNonNull(pattern, "pattern must not be null");
    }
  }

  /** Patterned fill patch carrying one foreground color. */
  record PatternForeground(ExcelFillPattern pattern, ColorInput foregroundColor)
      implements CellFillInput {
    public PatternForeground {
      requireForegroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
    }
  }

  /** Patterned fill patch carrying one background color. */
  record PatternBackground(ExcelFillPattern pattern, ColorInput backgroundColor)
      implements CellFillInput {
    public PatternBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Patterned fill patch carrying both foreground and background colors. */
  record PatternForegroundBackground(
      ExcelFillPattern pattern, ColorInput foregroundColor, ColorInput backgroundColor)
      implements CellFillInput {
    public PatternForegroundBackground {
      requireBackgroundPattern(pattern);
      Objects.requireNonNull(foregroundColor, "foregroundColor must not be null");
      Objects.requireNonNull(backgroundColor, "backgroundColor must not be null");
    }
  }

  /** Gradient fill patch. */
  record Gradient(CellGradientFillInput gradient) implements CellFillInput {
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
