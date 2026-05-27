package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.util.Map;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/** Converts between GridGrind style enums and the corresponding POI enums. */
final class ExcelCellStylePoiBridge {
  private static final Map<FillPatternType, ExcelFillPattern> FROM_POI_FILL_PATTERN =
      ExcelEnumMappingSupport.exactEnumMap(
          FillPatternType.class,
          "Excel cell fill-pattern mapping",
          Map.ofEntries(
              Map.entry(FillPatternType.NO_FILL, ExcelFillPattern.NONE),
              Map.entry(FillPatternType.SOLID_FOREGROUND, ExcelFillPattern.SOLID),
              Map.entry(FillPatternType.FINE_DOTS, ExcelFillPattern.FINE_DOTS),
              Map.entry(FillPatternType.ALT_BARS, ExcelFillPattern.ALT_BARS),
              Map.entry(FillPatternType.SPARSE_DOTS, ExcelFillPattern.SPARSE_DOTS),
              Map.entry(FillPatternType.THICK_HORZ_BANDS, ExcelFillPattern.THICK_HORIZONTAL_BANDS),
              Map.entry(FillPatternType.THICK_VERT_BANDS, ExcelFillPattern.THICK_VERTICAL_BANDS),
              Map.entry(
                  FillPatternType.THICK_BACKWARD_DIAG, ExcelFillPattern.THICK_BACKWARD_DIAGONAL),
              Map.entry(
                  FillPatternType.THICK_FORWARD_DIAG, ExcelFillPattern.THICK_FORWARD_DIAGONAL),
              Map.entry(FillPatternType.BIG_SPOTS, ExcelFillPattern.BIG_SPOTS),
              Map.entry(FillPatternType.BRICKS, ExcelFillPattern.BRICKS),
              Map.entry(FillPatternType.THIN_HORZ_BANDS, ExcelFillPattern.THIN_HORIZONTAL_BANDS),
              Map.entry(FillPatternType.THIN_VERT_BANDS, ExcelFillPattern.THIN_VERTICAL_BANDS),
              Map.entry(
                  FillPatternType.THIN_BACKWARD_DIAG, ExcelFillPattern.THIN_BACKWARD_DIAGONAL),
              Map.entry(FillPatternType.THIN_FORWARD_DIAG, ExcelFillPattern.THIN_FORWARD_DIAGONAL),
              Map.entry(FillPatternType.SQUARES, ExcelFillPattern.SQUARES),
              Map.entry(FillPatternType.DIAMONDS, ExcelFillPattern.DIAMONDS),
              Map.entry(FillPatternType.LESS_DOTS, ExcelFillPattern.LESS_DOTS),
              Map.entry(FillPatternType.LEAST_DOTS, ExcelFillPattern.LEAST_DOTS)));

  private static final Map<ExcelFillPattern, FillPatternType> TO_POI_FILL_PATTERN =
      ExcelEnumMappingSupport.reverseExactEnumMap(
          ExcelFillPattern.class, "Apache POI fill-pattern mapping", FROM_POI_FILL_PATTERN);

  private ExcelCellStylePoiBridge() {}

  static ExcelHorizontalAlignment fromPoi(HorizontalAlignment alignment) {
    return ExcelHorizontalAlignment.valueOf(alignment.name());
  }

  static ExcelVerticalAlignment fromPoi(VerticalAlignment alignment) {
    return ExcelVerticalAlignment.valueOf(alignment.name());
  }

  static ExcelBorderStyle fromPoi(BorderStyle borderStyle) {
    return ExcelBorderStyle.valueOf(borderStyle.name());
  }

  static HorizontalAlignment toPoi(ExcelHorizontalAlignment alignment) {
    return HorizontalAlignment.valueOf(alignment.name());
  }

  static VerticalAlignment toPoi(ExcelVerticalAlignment alignment) {
    return VerticalAlignment.valueOf(alignment.name());
  }

  static BorderStyle toPoi(ExcelBorderStyle borderStyle) {
    return BorderStyle.valueOf(borderStyle.name());
  }

  static ExcelFillPattern fromPoi(FillPatternType pattern) {
    return ExcelEnumMappingSupport.requireMappedValue(
        FROM_POI_FILL_PATTERN, pattern, "POI fill pattern");
  }

  static FillPatternType toPoi(ExcelFillPattern pattern) {
    return ExcelEnumMappingSupport.requireMappedValue(
        TO_POI_FILL_PATTERN, pattern, "GridGrind fill pattern");
  }
}
