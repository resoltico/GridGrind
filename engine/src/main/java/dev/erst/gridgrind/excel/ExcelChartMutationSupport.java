package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchorSupport;
import java.util.List;
import java.util.Objects;
import java.util.function.BooleanSupplier;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTx;

/** Chart authoring helpers. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelChartMutationSupport {
  private ExcelChartMutationSupport() {}

  public static void validateChart(XSSFSheet sheet, ResolvedChart definition) {
    validateChart(sheet, definition, null);
  }

  public static void validateChart(
      XSSFSheet sheet, ResolvedChart definition, @Nullable ExcelFormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    validateChartTitle(sheet, definition.title(), formulaRuntime);
    for (ResolvedChart.Plot plot : definition.plots()) {
      validatePlot(sheet, plot, formulaRuntime);
    }
  }

  public static void createChart(XSSFSheet sheet, ResolvedChart definition) {
    createChart(sheet, definition, null);
  }

  public static void createChart(
      XSSFSheet sheet, ResolvedChart definition, @Nullable ExcelFormulaRuntime formulaRuntime) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        ExcelDrawingAnchorSupport.toPoiAnchor(drawing, definition.anchor());
    XSSFChart chart = drawing.createChart(anchor);
    chart.getGraphicFrame().setName(definition.name());
    applyCommonChartState(sheet, chart, definition, formulaRuntime);

    ExcelChartAxisRegistry axisRegistry = new ExcelChartAxisRegistry(chart);
    for (ResolvedChart.Plot plot : definition.plots()) {
      ExcelChartPlotMutationSupport.createPlot(sheet, chart, axisRegistry, plot, formulaRuntime);
    }
  }

  static PreparedSeriesTitle prepareSeriesTitle(XSSFSheet sheet, ResolvedChartTitle title) {
    return prepareSeriesTitle(sheet, title, null);
  }

  static PreparedSeriesTitle prepareSeriesTitle(
      XSSFSheet sheet, ResolvedChartTitle title, @Nullable ExcelFormulaRuntime formulaRuntime) {
    return switch (title) {
      case ChartTitleSpec.None _ -> new PreparedSeriesTitleNone();
      case ResolvedChartTitle.Text text -> new PreparedSeriesTitleText(text.text());
      case ChartTitleSpec.Formula formula -> {
        CellReference reference =
            ExcelChartSourceSupport.resolveSingleCellReference(
                sheet,
                ExcelChartSourceSupport.normalizeFormula(formula.formula()),
                "Series title formula");
        yield new PreparedSeriesTitleFormula(
            ExcelChartSourceSupport.scalarText(sheet, reference, formulaRuntime), reference);
      }
    };
  }

  private static void validateChartTitle(
      XSSFSheet sheet, ResolvedChartTitle title, @Nullable ExcelFormulaRuntime formulaRuntime) {
    if (title instanceof ChartTitleSpec.Formula formula) {
      CellReference reference =
          ExcelChartSourceSupport.resolveSingleCellReference(
              sheet,
              ExcelChartSourceSupport.normalizeFormula(formula.formula()),
              "Chart title formula");
      ExcelChartSourceSupport.scalarText(sheet, reference, formulaRuntime);
    }
  }

  private static void validatePlot(
      XSSFSheet sheet, ResolvedChart.Plot plot, @Nullable ExcelFormulaRuntime formulaRuntime) {
    for (ResolvedChart.Series series : plotSeries(plot)) {
      ExcelChartSourceSupport.toCategoryDataSource(sheet, series.categories(), formulaRuntime);
      ExcelChartSourceSupport.toValueDataSource(sheet, series.values(), formulaRuntime);
      prepareSeriesTitle(sheet, series.title(), formulaRuntime);
    }
  }

  private static List<ResolvedChart.Series> plotSeries(ResolvedChart.Plot plot) {
    return switch (plot) {
      case ResolvedChart.Area area -> area.series();
      case ResolvedChart.Area3D area3D -> area3D.series();
      case ResolvedChart.Bar bar -> bar.series();
      case ResolvedChart.Bar3D bar3D -> bar3D.series();
      case ResolvedChart.Doughnut doughnut -> doughnut.series();
      case ResolvedChart.Line line -> line.series();
      case ResolvedChart.Line3D line3D -> line3D.series();
      case ResolvedChart.Pie pie -> pie.series();
      case ResolvedChart.Pie3D pie3D -> pie3D.series();
      case ResolvedChart.Radar radar -> radar.series();
      case ResolvedChart.Scatter scatter -> scatter.series();
      case ResolvedChart.Surface surface -> surface.series();
      case ResolvedChart.Surface3D surface3D -> surface3D.series();
    };
  }

  static void applySeriesTitle(
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series series,
      PreparedSeriesTitle title) {
    switch (title) {
      case PreparedSeriesTitleNone _ ->
          unsetSeriesTitleIfPresent(series.getCTBarSer()::isSetTx, series.getCTBarSer()::unsetTx);
      case PreparedSeriesTitleText text -> series.setTitle(text.text());
      case PreparedSeriesTitleFormula formula ->
          series.setTitle(formula.cachedText(), formula.reference());
    }
  }

  static void applySeriesTitle(
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series series,
      PreparedSeriesTitle title) {
    switch (title) {
      case PreparedSeriesTitleNone _ ->
          unsetSeriesTitleIfPresent(series.getCTLineSer()::isSetTx, series.getCTLineSer()::unsetTx);
      case PreparedSeriesTitleText text -> series.setTitle(text.text());
      case PreparedSeriesTitleFormula formula ->
          series.setTitle(formula.cachedText(), formula.reference());
    }
  }

  static void applySeriesTitle(
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series series,
      PreparedSeriesTitle title) {
    switch (title) {
      case PreparedSeriesTitleNone _ ->
          unsetSeriesTitleIfPresent(series.getCTPieSer()::isSetTx, series.getCTPieSer()::unsetTx);
      case PreparedSeriesTitleText text -> series.setTitle(text.text());
      case PreparedSeriesTitleFormula formula ->
          series.setTitle(formula.cachedText(), formula.reference());
    }
  }

  private static void applyCommonChartState(
      XSSFSheet sheet,
      XSSFChart chart,
      ResolvedChart definition,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    applyChartTitle(sheet, chart, definition.title(), formulaRuntime);
    applyChartLegend(chart, definition.legend());
    chart.displayBlanksAs(ExcelChartPoiBridge.toPoiDisplayBlanks(definition.displayBlanksAs()));
    chart.setPlotOnlyVisibleCells(definition.plotOnlyVisibleCells());
  }

  private static void applyChartTitle(
      XSSFSheet sheet,
      XSSFChart chart,
      ResolvedChartTitle title,
      @Nullable ExcelFormulaRuntime formulaRuntime) {
    switch (title) {
      case ChartTitleSpec.None _ -> chart.removeTitle();
      case ResolvedChartTitle.Text text -> chart.setTitleText(text.text());
      case ChartTitleSpec.Formula formula -> {
        CellReference reference =
            ExcelChartSourceSupport.resolveSingleCellReference(
                sheet,
                ExcelChartSourceSupport.normalizeFormula(formula.formula()),
                "Chart title formula");
        applyChartTitleFormula(
            chart, ExcelChartSourceSupport.scalarText(sheet, reference, formulaRuntime), reference);
      }
    }
  }

  static void applyChartTitleFormula(XSSFChart chart, String cachedText, CellReference reference) {
    CTTitle title =
        chart.getCTChart().isSetTitle()
            ? chart.getCTChart().getTitle()
            : chart.getCTChart().addNewTitle();
    CTTx text = title.isSetTx() ? title.getTx() : title.addNewTx();
    if (text.isSetRich()) {
      text.unsetRich();
    }
    CTStrRef strRef = text.isSetStrRef() ? text.getStrRef() : text.addNewStrRef();
    strRef.setF(reference.formatAsString());
    writeStringReferenceCache(strRef, cachedText);
  }

  private static void writeStringReferenceCache(CTStrRef reference, String cachedText) {
    if (reference.isSetStrCache()) {
      reference.unsetStrCache();
    }
    CTStrData cache = reference.addNewStrCache();
    cache.addNewPtCount().setVal(1);
    cache.addNewPt().setIdx(0);
    cache.getPtArray(0).setV(cachedText);
  }

  private static void applyChartLegend(XSSFChart chart, ChartLegendSpec legend) {
    switch (legend) {
      case ChartLegendSpec.Hidden _ -> chart.deleteLegend();
      case ChartLegendSpec.Visible visible ->
          chart
              .getOrAddLegend()
              .setPosition(ExcelChartPoiBridge.toPoiLegendPosition(visible.position()));
    }
  }

  private static void unsetSeriesTitleIfPresent(BooleanSupplier hasTitle, Runnable unsetTitle) {
    Objects.requireNonNull(hasTitle, "hasTitle must not be null");
    Objects.requireNonNull(unsetTitle, "unsetTitle must not be null");
    if (hasTitle.getAsBoolean()) {
      unsetTitle.run();
    }
  }
}
