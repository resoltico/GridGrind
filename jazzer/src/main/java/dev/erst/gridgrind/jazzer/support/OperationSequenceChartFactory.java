package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ChartAxisInput;
import dev.erst.gridgrind.contract.dto.ChartDataSourceInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartLegendInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
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
import java.util.Optional;

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

  private static ChartTitleInput nextChartTitleInput(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> new ChartTitleInput.None();
      case 1 -> new ChartTitleInput.Text(TextSourceInput.inline("Chart " + data.consumeInt(0, 9)));
      default -> new ChartTitleInput.Formula(data.consumeBoolean() ? "B1" : "C1");
    };
  }

  private static ExcelChartDefinition.Title nextExcelChartTitle(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> new ExcelChartDefinition.Title.None();
      case 1 -> new ExcelChartDefinition.Title.Text("Chart " + data.consumeInt(0, 9));
      default -> new ExcelChartDefinition.Title.Formula(data.consumeBoolean() ? "B1" : "C1");
    };
  }

  private static ChartLegendInput nextChartLegendInput(GridGrindFuzzData data) {
    return data.consumeBoolean()
        ? new ChartLegendInput.Hidden()
        : new ChartLegendInput.Visible(nextChartLegendPosition(data));
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

  static ExcelChartBarDirection nextChartBarDirection(GridGrindFuzzData data) {
    return data.consumeBoolean() ? ExcelChartBarDirection.COLUMN : ExcelChartBarDirection.BAR;
  }

  static ExcelChartGrouping nextChartGrouping(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> ExcelChartGrouping.STANDARD;
      case 1 -> ExcelChartGrouping.PERCENT_STACKED;
      default -> ExcelChartGrouping.STACKED;
    };
  }

  static ExcelChartBarGrouping nextChartBarGrouping(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> ExcelChartBarGrouping.CLUSTERED;
      case 1 -> ExcelChartBarGrouping.PERCENT_STACKED;
      default -> ExcelChartBarGrouping.STACKED;
    };
  }

  static ExcelChartBarShape nextChartBarShape(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 4) {
      case 0 -> ExcelChartBarShape.BOX;
      case 1 -> ExcelChartBarShape.CONE;
      case 2 -> ExcelChartBarShape.CONE_TO_MAX;
      default -> ExcelChartBarShape.CYLINDER;
    };
  }

  static ExcelChartRadarStyle nextChartRadarStyle(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 3) {
      case 0 -> ExcelChartRadarStyle.STANDARD;
      case 1 -> ExcelChartRadarStyle.MARKER;
      default -> ExcelChartRadarStyle.FILLED;
    };
  }

  static ExcelChartScatterStyle nextChartScatterStyle(GridGrindFuzzData data) {
    return switch (selectorSlot(nextSelectorByte(data)) % 6) {
      case 0 -> ExcelChartScatterStyle.LINE;
      case 1 -> ExcelChartScatterStyle.LINE_MARKER;
      case 2 -> ExcelChartScatterStyle.MARKER;
      case 3 -> ExcelChartScatterStyle.NONE;
      case 4 -> ExcelChartScatterStyle.SMOOTH;
      default -> ExcelChartScatterStyle.SMOOTH_MARKER;
    };
  }

  private static ChartPlotInput nextChartPlotInput(GridGrindFuzzData data) {
    return OperationSequenceChartPlotSupport.nextChartPlotInput(data);
  }

  private static ExcelChartDefinition.Plot nextExcelChartPlotDefinition(GridGrindFuzzData data) {
    return OperationSequenceChartPlotSupport.nextExcelChartPlotDefinition(data);
  }

  static Optional<Integer> nextOptionalInt(GridGrindFuzzData data, int minimum, int maximum) {
    return data.consumeBoolean()
        ? Optional.of(data.consumeInt(minimum, maximum))
        : Optional.empty();
  }

  static List<ChartAxisInput> nextChartAxesInputCategory() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  static List<ChartAxisInput> nextChartAxesInputScatter() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  static List<ChartAxisInput> nextChartAxesInputSurface() {
    return List.of(
        new ChartAxisInput(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartAxisInput(
            ExcelChartAxisKind.SERIES,
            ExcelChartAxisPosition.RIGHT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  static List<ExcelChartDefinition.Axis> nextExcelChartAxesCategory() {
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

  static List<ExcelChartDefinition.Axis> nextExcelChartAxesScatter() {
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

  static List<ExcelChartDefinition.Axis> nextExcelChartAxesSurface() {
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

  static List<ChartSeriesInput> nextChartSeriesInputs(GridGrindFuzzData data, boolean pieChart) {
    List<ChartSeriesInput> series = new ArrayList<>();
    series.add(nextChartSeriesInput(data, "B1", "B2:B4"));
    if (!pieChart && data.consumeBoolean()) {
      series.add(nextChartSeriesInput(data, "C1", "C2:C4"));
    }
    return List.copyOf(series);
  }

  static List<ExcelChartDefinition.Series> nextExcelChartSeries(
      GridGrindFuzzData data, boolean pieChart) {
    List<ExcelChartDefinition.Series> series = new ArrayList<>();
    series.add(nextExcelChartSeries(data, "B1", "B2:B4"));
    if (!pieChart && data.consumeBoolean()) {
      series.add(nextExcelChartSeries(data, "C1", "C2:C4"));
    }
    return List.copyOf(series);
  }

  private static ChartSeriesInput nextChartSeriesInput(
      GridGrindFuzzData data, String titleFormula, String valuesFormula) {
    return new ChartSeriesInput(
        data.consumeBoolean()
            ? new ChartTitleInput.Formula(titleFormula)
            : new ChartTitleInput.Text(TextSourceInput.inline("Series " + data.consumeInt(0, 9))),
        nextChartDataSourceInput(data, "A2:A4", data.consumeBoolean()),
        nextChartDataSourceInput(data, valuesFormula, true),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        data.consumeBoolean()
            ? Optional.of(Long.valueOf(data.consumeInt(0, 50)))
            : Optional.empty());
  }

  private static ExcelChartDefinition.Series nextExcelChartSeries(
      GridGrindFuzzData data, String titleFormula, String valuesFormula) {
    return new ExcelChartDefinition.Series(
        data.consumeBoolean()
            ? new ExcelChartDefinition.Title.Formula(titleFormula)
            : new ExcelChartDefinition.Title.Text("Series " + data.consumeInt(0, 9)),
        nextExcelChartDataSource(data, "A2:A4", data.consumeBoolean()),
        nextExcelChartDataSource(data, valuesFormula, true),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        data.consumeBoolean()
            ? Optional.of(Long.valueOf(data.consumeInt(0, 50)))
            : Optional.empty());
  }

  private static ChartDataSourceInput nextChartDataSourceInput(
      GridGrindFuzzData data, String formula, boolean numeric) {
    if (data.consumeBoolean()) {
      return new ChartDataSourceInput.Reference(formula);
    }
    if (numeric) {
      return new ChartDataSourceInput.NumericLiteral(nextNumericLiteralValues(data));
    }
    return new ChartDataSourceInput.StringLiteral(nextStringLiteralValues(data));
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

  static int nextSelectorByte(GridGrindFuzzData data) {
    return Byte.toUnsignedInt(data.consumeByte());
  }

  static int selectorSlot(int selector) {
    return selector & 0x0F;
  }
}
