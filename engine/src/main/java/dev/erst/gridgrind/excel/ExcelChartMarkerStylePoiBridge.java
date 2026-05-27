package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;

/** Maps chart marker-style enums and OOXML marker tokens between GridGrind and POI. */
final class ExcelChartMarkerStylePoiBridge {
  private static final Map<String, ExcelChartMarkerStyle> FROM_SYMBOL_TOKEN =
      Map.ofEntries(
          Map.entry("CIRCLE", ExcelChartMarkerStyle.CIRCLE),
          Map.entry("DASH", ExcelChartMarkerStyle.DASH),
          Map.entry("DIAMOND", ExcelChartMarkerStyle.DIAMOND),
          Map.entry("DOT", ExcelChartMarkerStyle.DOT),
          Map.entry("NONE", ExcelChartMarkerStyle.NONE),
          Map.entry("PICTURE", ExcelChartMarkerStyle.PICTURE),
          Map.entry("PLUS", ExcelChartMarkerStyle.PLUS),
          Map.entry("SQUARE", ExcelChartMarkerStyle.SQUARE),
          Map.entry("STAR", ExcelChartMarkerStyle.STAR),
          Map.entry("TRIANGLE", ExcelChartMarkerStyle.TRIANGLE),
          Map.entry("X", ExcelChartMarkerStyle.X));

  private ExcelChartMarkerStylePoiBridge() {}

  static ExcelChartMarkerStyle fromPoi(MarkerStyle style) {
    return ExcelChartMarkerStyle.valueOf(style.name());
  }

  static MarkerStyle toPoi(ExcelChartMarkerStyle style) {
    return MarkerStyle.valueOf(style.name());
  }

  static Optional<ExcelChartMarkerStyle> fromSymbolToken(String symbolToken) {
    if (symbolToken == null || symbolToken.isBlank()) {
      return Optional.empty();
    }
    return Optional.ofNullable(FROM_SYMBOL_TOKEN.get(symbolToken.toUpperCase(Locale.ROOT)));
  }
}
