package dev.erst.gridgrind.engine.styles;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

/** Reads and writes RGB, indexed and theme colors without flattening their semantics. */
public final class ColorAdapter {
  private ColorAdapter() {}

  /** Converts one authored color into the matching POI workbook color. */
  public static XSSFColor toXssfColor(XSSFWorkbook workbook, ColorSpec color) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(color, "color must not be null");

    XSSFColor xssfColor =
        switch (color) {
          case ColorSpec.Rgb rgb -> toXssfColor(workbook, rgb.rgb());
          case ColorSpec.Theme _ -> new XSSFColor(workbook.getStylesSource().getIndexedColors());
          case ColorSpec.Indexed _ -> new XSSFColor(workbook.getStylesSource().getIndexedColors());
        };
    switch (color) {
      case ColorSpec.Rgb rgb -> applyTint(xssfColor, rgb.tint());
      case ColorSpec.Theme theme -> {
        xssfColor.setTheme(theme.theme());
        applyTint(xssfColor, theme.tint());
      }
      case ColorSpec.Indexed indexed -> {
        xssfColor.setIndexed(indexed.indexed());
        applyTint(xssfColor, indexed.tint());
      }
    }
    return xssfColor;
  }

  /** Retains the factual RGB/theme/indexed payload when copying mutable workbook styles. */
  public static @Nullable ColorSpec copyOf(@Nullable CellColorReport color) {
    return color == null
        ? null
        : switch (color) {
          case CellColorReport.Rgb rgb -> new ColorSpec.Rgb(rgb.rgb(), rgb.tint());
          case CellColorReport.Theme theme -> new ColorSpec.Theme(theme.theme(), theme.tint());
          case CellColorReport.Indexed indexed ->
              new ColorSpec.Indexed(indexed.indexed(), indexed.tint());
        };
  }

  private static void applyTint(XSSFColor color, java.util.Optional<Double> tint) {
    tint.ifPresent(color::setTint);
  }

  /** Returns one factual snapshot for the supplied workbook color, or empty when absent. */
  public static Optional<CellColorReport> snapshot(@Nullable XSSFColor color) {
    if (color == null) {
      return Optional.empty();
    }
    String rgb = color.isRGB() ? toRgbHex(color).orElse(null) : null;
    Integer theme = color.isThemed() ? color.getTheme() : null;
    Integer indexed = color.isIndexed() ? Short.toUnsignedInt(color.getIndexed()) : null;
    Optional<Double> tint = color.hasTint() ? Optional.of(color.getTint()) : Optional.empty();
    if (rgb != null) {
      return Optional.of(new CellColorReport.Rgb(new RgbHex(rgb), tint));
    }
    if (theme != null) {
      return Optional.of(new CellColorReport.Theme(theme, tint));
    }
    if (indexed != null) {
      return Optional.of(new CellColorReport.Indexed(indexed, tint));
    }
    throw new IllegalStateException(
        "Workbook color could not be normalized: no RGB, theme, or indexed payload was present"
            + (color.isAuto() ? " (automatic color)." : "."));
  }

  /** Returns one factual snapshot for the supplied raw workbook color XML, or empty when absent. */
  public static Optional<CellColorReport> snapshot(XSSFWorkbook workbook, @Nullable CTColor color) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    if (color == null) {
      return Optional.empty();
    }
    XSSFColor xssfColor = XSSFColor.from(color, workbook.getStylesSource().getIndexedColors());
    ThemesTable themes = workbook.getStylesSource().getTheme();
    if (themes != null) {
      themes.inheritFromThemeAsRequired(xssfColor);
    }
    return snapshot(xssfColor);
  }

  /** Converts one POI workbook color into {@code #RRGGBB}, or empty when no RGB is available. */
  public static Optional<String> toRgbHex(XSSFColor color) {
    if (color == null) {
      return Optional.empty();
    }
    byte[] rgb = color.getRGB();
    if (rgb == null || rgb.length != 3) {
      return Optional.empty();
    }
    return Optional.of("#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
  }

  /** Converts one normalized {@code #RRGGBB} literal into the matching POI workbook color. */
  public static XSSFColor toXssfColor(XSSFWorkbook workbook, RgbHex color) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    String normalized = Objects.requireNonNull(color, "color must not be null").value();
    return new XSSFColor(
        new byte[] {
          (byte) Integer.parseInt(normalized.substring(1, 3), 16),
          (byte) Integer.parseInt(normalized.substring(3, 5), 16),
          (byte) Integer.parseInt(normalized.substring(5, 7), 16)
        },
        workbook.getStylesSource().getIndexedColors());
  }
}
