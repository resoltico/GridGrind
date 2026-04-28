package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarShape;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartRadarStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartScatterStyle;
import java.util.ArrayList;
import java.util.List;

/** Dedicated chart payload generator shared by protocol and engine-side Jazzer sequences. */
final class OperationSequenceChartFactory {
  private static final String DRAWING_CHART_NAME = "OpsChart";

  private OperationSequenceChartFactory() {}

  static ChartInput nextChartInput(GridGrindFuzzData data) {
    return new ChartInput(
        DRAWING_CHART_NAME,
        OperationSequenceValueFactory.nextDrawingAnchorInput(data),
        nextChartTitleInput(data),
        nextChartLegendInput(data),
        nextChartDisplayBlanksAs(data),
        data.consumeBoolean(),
        List.of(nextChartPlotInput(data)));
  }

  static ExcelChartDefinition nextExcelChartDefinition(GridGrindFuzzData data) {
    return new ExcelChartDefinition(
        DRAWING_CHART_NAME,
        OperationSequenceValueFactory.nextExcelDrawingAnchor(data),
        nextExcelChartTitle(data),
        nextExcelChartLegend(data),
        nextChartDisplayBlanksAs(data),
        data.consumeBoolean(),
        List.of(nextExcelChartPlotDefinition(data)));
  }

  private static ChartInput.Title nextChartTitleInput(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> new ChartInput.Title.None();
      case 1 -> new ChartInput.Title.Text(TextSourceInput.inline("Chart " + data.consumeInt(0, 9)));
      default -> new ChartInput.Title.Formula(data.consumeBoolean() ? "B1" : "C1");
    };
  }

  private static ExcelChartDefinition.Title nextExcelChartTitle(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> new ExcelChartDefinition.Title.None();
      case 1 -> new ExcelChartDefinition.Title.Text("Chart " + data.consumeInt(0, 9));
      default -> new ExcelChartDefinition.Title.Formula(data.consumeBoolean() ? "B1" : "C1");
    };
  }

  private static ChartInput.Legend nextChartLegendInput(GridGrindFuzzData data) {
    return data.consumeBoolean()
        ? new ChartInput.Legend.Hidden()
        : new ChartInput.Legend.Visible(nextChartLegendPosition(data));
  }

  private static ExcelChartDefinition.Legend nextExcelChartLegend(GridGrindFuzzData data) {
    return data.consumeBoolean()
        ? new ExcelChartDefinition.Legend.Hidden()
        : new ExcelChartDefinition.Legend.Visible(nextChartLegendPosition(data));
  }

  private static ExcelChartLegendPosition nextChartLegendPosition(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 5) {
      case 0 -> ExcelChartLegendPosition.BOTTOM;
      case 1 -> ExcelChartLegendPosition.LEFT;
      case 2 -> ExcelChartLegendPosition.RIGHT;
      case 3 -> ExcelChartLegendPosition.TOP;
      default -> ExcelChartLegendPosition.TOP_RIGHT;
    };
  }

  private static ExcelChartDisplayBlanksAs nextChartDisplayBlanksAs(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> ExcelChartDisplayBlanksAs.GAP;
      case 1 -> ExcelChartDisplayBlanksAs.SPAN;
      default -> ExcelChartDisplayBlanksAs.ZERO;
    };
  }

  private static ExcelChartBarDirection nextChartBarDirection(GridGrindFuzzData data) {
    return data.consumeBoolean() ? ExcelChartBarDirection.COLUMN : ExcelChartBarDirection.BAR;
  }

  private static ExcelChartGrouping nextChartGrouping(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> ExcelChartGrouping.STANDARD;
      case 1 -> ExcelChartGrouping.PERCENT_STACKED;
      default -> ExcelChartGrouping.STACKED;
    };
  }

  private static ExcelChartBarGrouping nextChartBarGrouping(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> ExcelChartBarGrouping.CLUSTERED;
      case 1 -> ExcelChartBarGrouping.PERCENT_STACKED;
      default -> ExcelChartBarGrouping.STACKED;
    };
  }

  private static ExcelChartBarShape nextChartBarShape(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 4) {
      case 0 -> ExcelChartBarShape.BOX;
      case 1 -> ExcelChartBarShape.CONE;
      case 2 -> ExcelChartBarShape.CONE_TO_MAX;
      default -> ExcelChartBarShape.CYLINDER;
    };
  }

  private static ExcelChartRadarStyle nextChartRadarStyle(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> ExcelChartRadarStyle.STANDARD;
      case 1 -> ExcelChartRadarStyle.MARKER;
      default -> ExcelChartRadarStyle.FILLED;
    };
  }

  private static ExcelChartScatterStyle nextChartScatterStyle(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 6) {
      case 0 -> ExcelChartScatterStyle.LINE;
      case 1 -> ExcelChartScatterStyle.LINE_MARKER;
      case 2 -> ExcelChartScatterStyle.MARKER;
      case 3 -> ExcelChartScatterStyle.NONE;
      case 4 -> ExcelChartScatterStyle.SMOOTH;
      default -> ExcelChartScatterStyle.SMOOTH_MARKER;
    };
  }

  private static ChartInput.Plot nextChartPlotInput(GridGrindFuzzData data) {
    Boolean varyColors = data.consumeBoolean();
    return switch (selectorSlot(nextSelectorByte(data)) % 13) {
      case 0 ->
          new ChartInput.Area(
              varyColors,
              nextChartGrouping(data),
              nextChartAxesInputCategory(),
              nextChartSeriesInputs(data, false));
      case 1 ->
          new ChartInput.Area3D(
              varyColors,
              nextChartGrouping(data),
              nextOptionalInt(data, 0, 500),
              nextChartAxesInputCategory(),
              nextChartSeriesInputs(data, false));
      case 2 ->
          new ChartInput.Bar(
              varyColors,
              nextChartBarDirection(data),
              nextChartBarGrouping(data),
              nextOptionalInt(data, 0, 500),
              nextOptionalInt(data, -100, 100),
              nextChartAxesInputCategory(),
              nextChartSeriesInputs(data, false));
      case 3 ->
          new ChartInput.Bar3D(
              varyColors,
              nextChartBarDirection(data),
              nextChartBarGrouping(data),
              nextOptionalInt(data, 0, 500),
              nextOptionalInt(data, 0, 500),
              nextChartBarShape(data),
              nextChartAxesInputCategory(),
              nextChartSeriesInputs(data, false));
      case 4 ->
          new ChartInput.Doughnut(
              varyColors,
              nextOptionalInt(data, 0, 360),
              nextOptionalInt(data, 10, 90),
              nextChartSeriesInputs(data, true));
      case 5 ->
          new ChartInput.Line(
              varyColors,
              nextChartGrouping(data),
              nextChartAxesInputCategory(),
              nextChartSeriesInputs(data, false));
      case 6 ->
          new ChartInput.Line3D(
              varyColors,
              nextChartGrouping(data),
              nextOptionalInt(data, 0, 500),
              nextChartAxesInputCategory(),
              nextChartSeriesInputs(data, false));
      case 7 ->
          new ChartInput.Pie(
              varyColors, nextOptionalInt(data, 0, 360), nextChartSeriesInputs(data, true));
      case 8 -> new ChartInput.Pie3D(varyColors, nextChartSeriesInputs(data, true));
      case 9 ->
          new ChartInput.Radar(
              varyColors,
              nextChartRadarStyle(data),
              nextChartAxesInputCategory(),
              nextChartSeriesInputs(data, false));
      case 10 ->
          new ChartInput.Scatter(
              varyColors,
              nextChartScatterStyle(data),
              nextChartAxesInputScatter(),
              nextChartSeriesInputs(data, false));
      case 11 ->
          new ChartInput.Surface(
              varyColors,
              data.consumeBoolean(),
              nextChartAxesInputSurface(),
              nextChartSeriesInputs(data, false));
      default ->
          new ChartInput.Surface3D(
              varyColors,
              data.consumeBoolean(),
              nextChartAxesInputSurface(),
              nextChartSeriesInputs(data, false));
    };
  }

  private static ExcelChartDefinition.Plot nextExcelChartPlotDefinition(GridGrindFuzzData data) {
    boolean varyColors = data.consumeBoolean();
    return switch (selectorSlot(nextSelectorByte(data)) % 13) {
      case 0 ->
          new ExcelChartDefinition.Area(
              varyColors,
              nextChartGrouping(data),
              nextExcelChartAxesCategory(),
              nextExcelChartSeries(data, false));
      case 1 ->
          new ExcelChartDefinition.Area3D(
              varyColors,
              nextChartGrouping(data),
              nextOptionalInt(data, 0, 500),
              nextExcelChartAxesCategory(),
              nextExcelChartSeries(data, false));
      case 2 ->
          new ExcelChartDefinition.Bar(
              varyColors,
              nextChartBarDirection(data),
              nextChartBarGrouping(data),
              nextOptionalInt(data, 0, 500),
              nextOptionalInt(data, -100, 100),
              nextExcelChartAxesCategory(),
              nextExcelChartSeries(data, false));
      case 3 ->
          new ExcelChartDefinition.Bar3D(
              varyColors,
              nextChartBarDirection(data),
              nextChartBarGrouping(data),
              nextOptionalInt(data, 0, 500),
              nextOptionalInt(data, 0, 500),
              nextChartBarShape(data),
              nextExcelChartAxesCategory(),
              nextExcelChartSeries(data, false));
      case 4 ->
          new ExcelChartDefinition.Doughnut(
              varyColors,
              nextOptionalInt(data, 0, 360),
              nextOptionalInt(data, 10, 90),
              nextExcelChartSeries(data, true));
      case 5 ->
          new ExcelChartDefinition.Line(
              varyColors,
              nextChartGrouping(data),
              nextExcelChartAxesCategory(),
              nextExcelChartSeries(data, false));
      case 6 ->
          new ExcelChartDefinition.Line3D(
              varyColors,
              nextChartGrouping(data),
              nextOptionalInt(data, 0, 500),
              nextExcelChartAxesCategory(),
              nextExcelChartSeries(data, false));
      case 7 ->
          new ExcelChartDefinition.Pie(
              varyColors, nextOptionalInt(data, 0, 360), nextExcelChartSeries(data, true));
      case 8 -> new ExcelChartDefinition.Pie3D(varyColors, nextExcelChartSeries(data, true));
      case 9 ->
          new ExcelChartDefinition.Radar(
              varyColors,
              nextChartRadarStyle(data),
              nextExcelChartAxesCategory(),
              nextExcelChartSeries(data, false));
      case 10 ->
          new ExcelChartDefinition.Scatter(
              varyColors,
              nextChartScatterStyle(data),
              nextExcelChartAxesScatter(),
              nextExcelChartSeries(data, false));
      case 11 ->
          new ExcelChartDefinition.Surface(
              varyColors,
              data.consumeBoolean(),
              nextExcelChartAxesSurface(),
              nextExcelChartSeries(data, false));
      default ->
          new ExcelChartDefinition.Surface3D(
              varyColors,
              data.consumeBoolean(),
              nextExcelChartAxesSurface(),
              nextExcelChartSeries(data, false));
    };
  }

  private static Integer nextOptionalInt(GridGrindFuzzData data, int minimum, int maximum) {
    return data.consumeBoolean() ? data.consumeInt(minimum, maximum) : null;
  }

  private static List<ChartInput.Axis> nextChartAxesInputCategory() {
    return List.of(
        new ChartInput.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartInput.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartInput.Axis> nextChartAxesInputScatter() {
    return List.of(
        new ChartInput.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartInput.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartInput.Axis> nextChartAxesInputSurface() {
    return List.of(
        new ChartInput.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartInput.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartInput.Axis(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartDefinition.Axis> nextExcelChartAxesCategory() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartDefinition.Axis> nextExcelChartAxesScatter() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartDefinition.Axis> nextExcelChartAxesSurface() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ChartInput.Series> nextChartSeriesInputs(
      GridGrindFuzzData data, boolean pieChart) {
    List<ChartInput.Series> series = new ArrayList<>();
    series.add(nextChartSeriesInput(data, "B1", "B2:B4"));
    if (!pieChart && data.consumeBoolean()) {
      series.add(nextChartSeriesInput(data, "C1", "C2:C4"));
    }
    return List.copyOf(series);
  }

  private static List<ExcelChartDefinition.Series> nextExcelChartSeries(
      GridGrindFuzzData data, boolean pieChart) {
    List<ExcelChartDefinition.Series> series = new ArrayList<>();
    series.add(nextExcelChartSeries(data, "B1", "B2:B4"));
    if (!pieChart && data.consumeBoolean()) {
      series.add(nextExcelChartSeries(data, "C1", "C2:C4"));
    }
    return List.copyOf(series);
  }

  private static ChartInput.Series nextChartSeriesInput(
      GridGrindFuzzData data, String titleFormula, String valuesFormula) {
    return new ChartInput.Series(
        data.consumeBoolean()
            ? new ChartInput.Title.Formula(titleFormula)
            : new ChartInput.Title.Text(TextSourceInput.inline("Series " + data.consumeInt(0, 9))),
        nextChartDataSourceInput(data, "A2:A4", data.consumeBoolean()),
        nextChartDataSourceInput(data, valuesFormula, true),
        null,
        null,
        null,
        data.consumeBoolean() ? Long.valueOf(data.consumeInt(0, 50)) : null);
  }

  private static ExcelChartDefinition.Series nextExcelChartSeries(
      GridGrindFuzzData data, String titleFormula, String valuesFormula) {
    return new ExcelChartDefinition.Series(
        data.consumeBoolean()
            ? new ExcelChartDefinition.Title.Formula(titleFormula)
            : new ExcelChartDefinition.Title.Text("Series " + data.consumeInt(0, 9)),
        nextExcelChartDataSource(data, "A2:A4", data.consumeBoolean()),
        nextExcelChartDataSource(data, valuesFormula, true),
        null,
        null,
        null,
        data.consumeBoolean() ? Long.valueOf(data.consumeInt(0, 50)) : null);
  }

  private static ChartInput.DataSource nextChartDataSourceInput(
      GridGrindFuzzData data, String formula, boolean numeric) {
    if (data.consumeBoolean()) {
      return new ChartInput.DataSource.Reference(formula);
    }
    if (numeric) {
      return new ChartInput.DataSource.NumericLiteral(nextNumericLiteralValues(data));
    }
    return new ChartInput.DataSource.StringLiteral(nextStringLiteralValues(data));
  }

  private static ExcelChartDefinition.DataSource nextExcelChartDataSource(
      GridGrindFuzzData data, String formula, boolean numeric) {
    if (data.consumeBoolean()) {
      return new ExcelChartDefinition.DataSource.Reference(formula);
    }
    if (numeric) {
      return new ExcelChartDefinition.DataSource.NumericLiteral(nextNumericLiteralValues(data));
    }
    return new ExcelChartDefinition.DataSource.StringLiteral(nextStringLiteralValues(data));
  }

  private static List<Double> nextNumericLiteralValues(GridGrindFuzzData data) {
    return List.of(
        (double) data.consumeInt(1, 9),
        (double) data.consumeInt(10, 19),
        (double) data.consumeInt(20, 29));
  }

  private static List<String> nextStringLiteralValues(GridGrindFuzzData data) {
    return List.of(
        "Label " + data.consumeInt(0, 9),
        "Label " + data.consumeInt(10, 19),
        "Label " + data.consumeInt(20, 29));
  }

  private static int nextSelectorByte(GridGrindFuzzData data) {
    return Byte.toUnsignedInt(data.consumeByte());
  }

  private static int selectorSlot(int selector) {
    return selector & 0x0F;
  }
}
