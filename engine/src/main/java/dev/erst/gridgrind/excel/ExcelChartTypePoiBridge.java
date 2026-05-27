package dev.erst.gridgrind.excel;

import java.util.Map;
import java.util.Objects;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;

/** Maps GridGrind chart-plot kinds to the Apache POI chart data family enum. */
final class ExcelChartTypePoiBridge {
  private static final Map<ExcelChartPlotType, ChartTypes> TO_POI_CHART_TYPE =
      ExcelEnumMappingSupport.exactEnumMap(
          ExcelChartPlotType.class,
          "Apache POI chart type mapping",
          Map.ofEntries(
              Map.entry(ExcelChartPlotType.AREA, ChartTypes.AREA),
              Map.entry(ExcelChartPlotType.AREA_3D, ChartTypes.AREA3D),
              Map.entry(ExcelChartPlotType.BAR, ChartTypes.BAR),
              Map.entry(ExcelChartPlotType.BAR_3D, ChartTypes.BAR3D),
              Map.entry(ExcelChartPlotType.DOUGHNUT, ChartTypes.DOUGHNUT),
              Map.entry(ExcelChartPlotType.LINE, ChartTypes.LINE),
              Map.entry(ExcelChartPlotType.LINE_3D, ChartTypes.LINE3D),
              Map.entry(ExcelChartPlotType.PIE, ChartTypes.PIE),
              Map.entry(ExcelChartPlotType.PIE_3D, ChartTypes.PIE3D),
              Map.entry(ExcelChartPlotType.RADAR, ChartTypes.RADAR),
              Map.entry(ExcelChartPlotType.SCATTER, ChartTypes.SCATTER),
              Map.entry(ExcelChartPlotType.SURFACE, ChartTypes.SURFACE),
              Map.entry(ExcelChartPlotType.SURFACE_3D, ChartTypes.SURFACE3D)));

  private ExcelChartTypePoiBridge() {}

  static ChartTypes toPoiChartType(ExcelChartPlotType plotType) {
    Objects.requireNonNull(plotType, "plotType must not be null");
    return ExcelEnumMappingSupport.requireMappedValue(
        TO_POI_CHART_TYPE, plotType, "GridGrind chart plot type");
  }
}
