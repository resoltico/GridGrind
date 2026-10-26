package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.Objects;
import java.util.Optional;

/** Factual workbook color preserving RGB, theme, indexed, and tint semantics. */
@Union(discriminator = "kind", catalog = true, displayOrder = 58)
public sealed interface CellColorReport
    permits dev.erst.gridgrind.core.model.style.CellColorReport.Rgb,
        dev.erst.gridgrind.core.model.style.CellColorReport.Theme,
        dev.erst.gridgrind.core.model.style.CellColorReport.Indexed {
  /** Optional tint adjustment applied to the base color reference. */
  Optional<Double> tint();

  /** Returns one color report carrying only explicit RGB data. */
  static Rgb rgb(RgbHex rgb) {
    return new Rgb(rgb, Optional.empty());
  }

  /** Returns one RGB-backed report plus tint metadata. */
  static Rgb rgb(RgbHex rgb, double tint) {
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
  @Variant(
      id = "RGB",
      summary = "Read one explicit RGB workbook color plus optional tint metadata.")
  record Rgb(RgbHex rgb, Optional<Double> tint) implements CellColorReport {
    public Rgb {
      java.util.Objects.requireNonNull(rgb, "rgb must not be null");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Theme-backed workbook color report. */
  @Variant(
      id = "THEME",
      summary = "Read one workbook theme-slot color plus optional tint metadata.")
  record Theme(int theme, Optional<Double> tint) implements CellColorReport {
    public Theme {
      requireNonNegative(theme, "theme");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Indexed-palette workbook color report. */
  @Variant(
      id = "INDEXED",
      summary = "Read one indexed-palette workbook color plus optional tint metadata.")
  record Indexed(int indexed, Optional<Double> tint) implements CellColorReport {
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
