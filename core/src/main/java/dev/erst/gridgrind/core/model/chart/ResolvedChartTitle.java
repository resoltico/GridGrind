package dev.erst.gridgrind.core.model.chart;

/** Materialized chart titles, sharing source-free authored alternatives. */
public sealed interface ResolvedChartTitle
    permits ChartTitleSpec.None, ChartTitleSpec.Formula, ResolvedChartTitle.Text {
  /** Static title text after source materialization. */
  record Text(String text) implements ResolvedChartTitle {
    public Text {
      text = ChartSpec.requireNonBlank(text, "text");
    }
  }
}
