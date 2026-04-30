package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Factual workbook color preserving RGB, theme, indexed, and tint semantics. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellColorReport.Rgb.class, name = "RGB"),
  @JsonSubTypes.Type(value = CellColorReport.Theme.class, name = "THEME"),
  @JsonSubTypes.Type(value = CellColorReport.Indexed.class, name = "INDEXED")
})
public sealed interface CellColorReport
    permits CellColorReport.Rgb, CellColorReport.Theme, CellColorReport.Indexed {
  /** Optional tint adjustment applied to the base color reference. */
  Double tint();

  /** Returns one color report carrying only explicit RGB data. */
  static Rgb rgb(String rgb) {
    return new Rgb(rgb, null);
  }

  /** Returns one RGB-backed report plus tint metadata. */
  static Rgb rgb(String rgb, Double tint) {
    return new Rgb(rgb, tint);
  }

  /** Returns one theme-backed report with no tint adjustment. */
  static Theme theme(int theme) {
    return new Theme(theme, null);
  }

  /** Returns one theme-backed report plus tint metadata. */
  static Theme theme(int theme, Double tint) {
    return new Theme(theme, tint);
  }

  /** Returns one indexed-palette report with no tint adjustment. */
  static Indexed indexed(int indexed) {
    return new Indexed(indexed, null);
  }

  /** Returns one indexed-palette report plus tint metadata. */
  static Indexed indexed(int indexed, Double tint) {
    return new Indexed(indexed, tint);
  }

  /** RGB-backed workbook color report. */
  record Rgb(String rgb, Double tint) implements CellColorReport {
    public Rgb {
      rgb = ProtocolRgbColorSupport.requireRgbHex(rgb, "rgb");
      requireFiniteOrNull(tint, "tint");
    }
  }

  /** Theme-backed workbook color report. */
  record Theme(int theme, Double tint) implements CellColorReport {
    public Theme {
      requireNonNegative(theme, "theme");
      requireFiniteOrNull(tint, "tint");
    }
  }

  /** Indexed-palette workbook color report. */
  record Indexed(int indexed, Double tint) implements CellColorReport {
    public Indexed {
      requireNonNegative(indexed, "indexed");
      requireFiniteOrNull(tint, "tint");
    }
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }

  private static void requireFiniteOrNull(Double value, String fieldName) {
    if (value != null && !Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
  }
}
