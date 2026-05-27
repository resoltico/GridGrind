package dev.erst.gridgrind.excel;

import java.util.Locale;
import java.util.Objects;
import org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFSurface3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFSurfaceChartData;

/** Maps POI chart-data families to GridGrind plot kinds and canonical tokens. */
final class ExcelChartDataFamilyPoiBridge {
  private ExcelChartDataFamilyPoiBridge() {}

  static ExcelChartPlotType plotType(XDDFChartData chartData) {
    Objects.requireNonNull(chartData, "chartData must not be null");
    return switch (chartData) {
      case XDDFAreaChartData _ -> ExcelChartPlotType.AREA;
      case XDDFArea3DChartData _ -> ExcelChartPlotType.AREA_3D;
      case XDDFBarChartData _ -> ExcelChartPlotType.BAR;
      case XDDFBar3DChartData _ -> ExcelChartPlotType.BAR_3D;
      case XDDFDoughnutChartData _ -> ExcelChartPlotType.DOUGHNUT;
      case XDDFLineChartData _ -> ExcelChartPlotType.LINE;
      case XDDFLine3DChartData _ -> ExcelChartPlotType.LINE_3D;
      case XDDFPieChartData _ -> ExcelChartPlotType.PIE;
      case XDDFPie3DChartData _ -> ExcelChartPlotType.PIE_3D;
      case XDDFRadarChartData _ -> ExcelChartPlotType.RADAR;
      case XDDFScatterChartData _ -> ExcelChartPlotType.SCATTER;
      case XDDFSurfaceChartData _ -> ExcelChartPlotType.SURFACE;
      case XDDFSurface3DChartData _ -> ExcelChartPlotType.SURFACE_3D;
      default -> throw new IllegalArgumentException("Unsupported chart data family: " + chartData);
    };
  }

  static String plotTypeToken(XDDFChartData chartData) {
    Objects.requireNonNull(chartData, "chartData must not be null");
    return switch (chartData) {
      case XDDFAreaChartData _ -> "AREA";
      case XDDFArea3DChartData _ -> "AREA_3D";
      case XDDFBarChartData _ -> "BAR";
      case XDDFBar3DChartData _ -> "BAR_3D";
      case XDDFDoughnutChartData _ -> "DOUGHNUT";
      case XDDFLineChartData _ -> "LINE";
      case XDDFLine3DChartData _ -> "LINE_3D";
      case XDDFPieChartData _ -> "PIE";
      case XDDFPie3DChartData _ -> "PIE_3D";
      case XDDFRadarChartData _ -> "RADAR";
      case XDDFScatterChartData _ -> "SCATTER";
      case XDDFSurfaceChartData _ -> "SURFACE";
      case XDDFSurface3DChartData _ -> "SURFACE_3D";
      default -> canonicalPlotTypeToken(chartData.getClass().getSimpleName());
    };
  }

  static String canonicalPlotTypeToken(String simpleName) {
    Objects.requireNonNull(simpleName, "simpleName must not be null");
    if (simpleName.startsWith("XDDF") && simpleName.endsWith("ChartData")) {
      return simpleName
          .substring(4, simpleName.length() - "ChartData".length())
          .toUpperCase(Locale.ROOT);
    }
    return simpleName.toUpperCase(Locale.ROOT);
  }
}
