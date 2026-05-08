package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;
import java.util.Optional;

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
  Optional<Double> tint();

  /** Returns one color report carrying only explicit RGB data. */
  static Rgb rgb(String rgb) {
    return new Rgb(rgb, Optional.empty());
  }

  /** Returns one RGB-backed report plus tint metadata. */
  static Rgb rgb(String rgb, double tint) {
    return new Rgb(rgb, Optional.of(tint));
  }

  /** Returns one theme-backed report with no tint adjustment. */
  static Theme theme(int theme) {
    return new Theme(theme, Optional.empty());
  }

  /** Returns one theme-backed report plus tint metadata. */
  static Theme theme(int theme, double tint) {
    return new Theme(theme, Optional.of(tint));
  }

  /** Returns one indexed-palette report with no tint adjustment. */
  static Indexed indexed(int indexed) {
    return new Indexed(indexed, Optional.empty());
  }

  /** Returns one indexed-palette report plus tint metadata. */
  static Indexed indexed(int indexed, double tint) {
    return new Indexed(indexed, Optional.of(tint));
  }

  /** RGB-backed workbook color report. */
  record Rgb(String rgb, @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> tint)
      implements CellColorReport {
    public Rgb {
      rgb = ProtocolRgbColorSupport.requireRgbHex(rgb, "rgb");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Theme-backed workbook color report. */
  record Theme(int theme, @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> tint)
      implements CellColorReport {
    public Theme {
      requireNonNegative(theme, "theme");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Indexed-palette workbook color report. */
  record Indexed(int indexed, @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> tint)
      implements CellColorReport {
    public Indexed {
      requireNonNegative(indexed, "indexed");
      tint = requireFinite(tint, "tint");
    }
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }

  private static Optional<Double> requireFinite(Optional<Double> value, String fieldName) {
    Optional<Double> normalized = Objects.requireNonNull(value, fieldName + " must not be null");
    if (normalized.isPresent() && !Double.isFinite(normalized.orElseThrow())) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    return normalized;
  }
}
