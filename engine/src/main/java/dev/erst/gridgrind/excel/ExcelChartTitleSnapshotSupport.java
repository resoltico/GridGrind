package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

/** Title extraction and formula-resolution helpers for chart snapshots. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelChartTitleSnapshotSupport {
  private static final System.Logger LOGGER =
      System.getLogger(ExcelChartTitleSnapshotSupport.class.getName());

  private ExcelChartTitleSnapshotSupport() {}

  public static ChartReport.Title snapshotTitle(XSSFChart chart) {
    return snapshotTitle(chart, null);
  }

  static ChartReport.Title snapshotTitle(
      XSSFChart chart, @Nullable ExcelFormulaRuntime formulaRuntime) {
    return snapshotTitle(chart, chart.getGraphicFrame(), formulaRuntime);
  }

  static ChartReport.Title snapshotTitle(
      XSSFChart chart,
      @Nullable XSSFGraphicFrame graphicFrame,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    if (!chart.getCTChart().isSetTitle()) {
      return new ChartReport.Title.None();
    }
    String formula = chart.getTitleFormula();
    if (formula != null) {
      return new ChartReport.Title.Formula(
          formula, cachedTitleText(chart, graphicFrame, formula, formulaRuntime));
    }
    String text = chart.getTitleText().getString();
    return text.isBlank() ? new ChartReport.Title.None() : new ChartReport.Title.Text(text);
  }

  public static String cachedTitleText(XSSFChart chart, String formula) {
    return cachedTitleText(chart, formula, null);
  }

  static String cachedTitleText(
      XSSFChart chart, String formula, @Nullable ExcelFormulaRuntime formulaRuntime) {
    return cachedTitleText(chart, chart.getGraphicFrame(), formula, formulaRuntime);
  }

  static String cachedTitleText(
      XSSFChart chart,
      @Nullable XSSFGraphicFrame graphicFrame,
      String formula,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    Optional<String> resolvedText =
        optionalResolvedTitleFormulaText(chart, graphicFrame, formula, formulaRuntime);
    if (resolvedText.isPresent()) {
      return resolvedText.orElseThrow();
    }
    if (!chart.getCTChart().isSetTitle()
        || !chart.getCTChart().getTitle().isSetTx()
        || !chart.getCTChart().getTitle().getTx().isSetStrRef()
        || !chart.getCTChart().getTitle().getTx().getStrRef().isSetStrCache()
        || chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().sizeOfPtArray() == 0) {
      return "";
    }
    return chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().getPtArray(0).getV();
  }

  public static String resolvedTitleFormulaText(XSSFChart chart, String formula) {
    return resolvedTitleFormulaText(chart, formula, null);
  }

  static String resolvedTitleFormulaText(
      XSSFChart chart, String formula, @Nullable ExcelFormulaRuntime formulaRuntime) {
    return optionalResolvedTitleFormulaText(chart, chart.getGraphicFrame(), formula, formulaRuntime)
        .orElse("");
  }

  static Optional<String> optionalResolvedTitleFormulaText(
      XSSFChart chart,
      @Nullable XSSFGraphicFrame graphicFrame,
      String formula,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    try {
      XSSFSheet contextSheet = ExcelChartRelationSupport.contextSheet(chart, graphicFrame);
      return contextSheet == null
          ? Optional.empty()
          : Optional.of(
              ExcelChartSourceSupport.scalarText(
                  contextSheet,
                  ExcelChartSourceSupport.resolveSingleCellReference(
                      contextSheet, formula, "Chart title formula"),
                  formulaRuntime));
    } catch (IllegalArgumentException exception) {
      if (recoverableTitleFormulaResolutionFailure(exception)) {
        return Optional.empty();
      }
      LOGGER.log(
          System.Logger.Level.WARNING,
          "Failed to resolve chart title formula '" + formula + "'; using cached or empty title",
          exception);
      return Optional.empty();
    }
  }

  public static ChartReport.Title snapshotSeriesTitle(CTSerTx title) {
    return snapshotSeriesTitle(null, title, null);
  }

  public static ChartReport.Title snapshotSeriesTitle(
      @Nullable XSSFSheet contextSheet,
      @Nullable CTSerTx title,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    if (title == null) {
      return new ChartReport.Title.None();
    }
    if (title.isSetStrRef()) {
      if (contextSheet != null) {
        try {
          return new ChartReport.Title.Formula(
              title.getStrRef().getF(),
              ExcelChartSourceSupport.scalarText(
                  contextSheet,
                  ExcelChartSourceSupport.resolveSingleCellReference(
                      contextSheet, title.getStrRef().getF(), "Series title formula"),
                  formulaRuntime));
        } catch (IllegalArgumentException ignored) {
          // Fall back to the embedded chart cache when the formula cannot be resolved live.
        }
      }
      String cachedText =
          title.getStrRef().isSetStrCache() && title.getStrRef().getStrCache().sizeOfPtArray() > 0
              ? title.getStrRef().getStrCache().getPtArray(0).getV()
              : "";
      return new ChartReport.Title.Formula(title.getStrRef().getF(), cachedText);
    }
    return title.isSetV() ? new ChartReport.Title.Text(title.getV()) : new ChartReport.Title.None();
  }

  public static String titleSummary(ChartReport.Title title) {
    return switch (title) {
      case ChartReport.Title.None _ -> "";
      case ChartReport.Title.Text text -> text.text();
      case ChartReport.Title.Formula formula ->
          formula.cachedText().isEmpty() ? formula.formula() : formula.cachedText();
    };
  }

  static String resolvedChartName(XSSFGraphicFrame graphicFrame) {
    String name = ExcelChartSourceSupport.blankAsOptional(graphicFrame.getName()).orElse(null);
    return name != null ? name : "Chart-" + graphicFrame.getId();
  }

  private static boolean recoverableTitleFormulaResolutionFailure(
      IllegalArgumentException exception) {
    return "Chart source formulas must not cache error values".equals(exception.getMessage());
  }
}
