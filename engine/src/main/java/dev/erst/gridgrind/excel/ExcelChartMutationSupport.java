package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.BooleanSupplier;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Chart authoring and mutation helpers. */
final class ExcelChartMutationSupport {
  private ExcelChartMutationSupport() {}

  static PreparedChartDefinition prepareChartDefinition(
      XSSFSheet sheet, ExcelChartDefinition definition) {
    PreparedChartTitle title = prepareChartTitle(sheet, definition.title());
    List<PreparedChartSeries> series = prepareChartSeries(sheet, definition.series());
    return switch (definition) {
      case ExcelChartDefinition.Bar bar ->
          new PreparedBarChart(
              bar.name(),
              bar.anchor(),
              title,
              bar.legend(),
              bar.displayBlanksAs(),
              bar.plotOnlyVisibleCells(),
              bar.varyColors(),
              bar.barDirection(),
              series);
      case ExcelChartDefinition.Line line ->
          new PreparedLineChart(
              line.name(),
              line.anchor(),
              title,
              line.legend(),
              line.displayBlanksAs(),
              line.plotOnlyVisibleCells(),
              line.varyColors(),
              series);
      case ExcelChartDefinition.Pie pie ->
          new PreparedPieChart(
              pie.name(),
              pie.anchor(),
              title,
              pie.legend(),
              pie.displayBlanksAs(),
              pie.plotOnlyVisibleCells(),
              pie.varyColors(),
              pie.firstSliceAngle(),
              series);
    };
  }

  static void createChart(XSSFSheet sheet, PreparedChartDefinition definition) {
    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
        ExcelDrawingAnchorSupport.toPoiAnchor(drawing, definition.anchor());
    XSSFChart chart = drawing.createChart(anchor);
    chart.getGraphicFrame().setName(definition.name());
    applyCommonChartState(chart, definition);
    switch (definition) {
      case PreparedBarChart bar -> initializeBarChart(chart, bar);
      case PreparedLineChart line -> initializeLineChart(chart, line);
      case PreparedPieChart pie -> initializePieChart(chart, pie);
    }
  }

  static void mutateChart(
      XSSFSheet sheet,
      ExcelDrawingController.LocatedShape located,
      XSSFChart chart,
      PreparedChartDefinition definition) {
    ExcelChartSnapshot snapshot =
        ExcelChartSnapshotSupport.snapshotChart(
            chart, (org.apache.poi.xssf.usermodel.XSSFGraphicFrame) located.shape());
    if (snapshot instanceof ExcelChartSnapshot.Unsupported unsupported) {
      throw new IllegalArgumentException(
          "Chart '"
              + unsupported.name()
              + "' on sheet '"
              + sheet.getSheetName()
              + "' contains unsupported detail and cannot be mutated authoritatively: "
              + unsupported.detail());
    }

    switch ((ExcelChartSnapshot.Supported) snapshot) {
      case ExcelChartSnapshot.Bar _ -> {
        if (!(definition instanceof PreparedBarChart bar)) {
          throw typeChangeNotSupported(sheet, definition.name());
        }
        ExcelDrawingAnchorSupport.updateAnchorInPlace(
            sheet, definition.name(), located.parentAnchor(), definition.anchor());
        applyCommonChartState(chart, definition);
        mutateBarChart(chart, bar);
      }
      case ExcelChartSnapshot.Line _ -> {
        if (!(definition instanceof PreparedLineChart line)) {
          throw typeChangeNotSupported(sheet, definition.name());
        }
        ExcelDrawingAnchorSupport.updateAnchorInPlace(
            sheet, definition.name(), located.parentAnchor(), definition.anchor());
        applyCommonChartState(chart, definition);
        mutateLineChart(chart, line);
      }
      case ExcelChartSnapshot.Pie _ -> {
        if (!(definition instanceof PreparedPieChart pie)) {
          throw typeChangeNotSupported(sheet, definition.name());
        }
        ExcelDrawingAnchorSupport.updateAnchorInPlace(
            sheet, definition.name(), located.parentAnchor(), definition.anchor());
        applyCommonChartState(chart, definition);
        mutatePieChart(chart, pie);
      }
    }
  }

  static PreparedSeriesTitle prepareSeriesTitle(XSSFSheet sheet, ExcelChartDefinition.Title title) {
    return switch (title) {
      case ExcelChartDefinition.Title.None _ -> new PreparedSeriesTitleNone();
      case ExcelChartDefinition.Title.Text text -> new PreparedSeriesTitleText(text.text());
      case ExcelChartDefinition.Title.Formula formula -> {
        CellReference reference =
            ExcelChartSourceSupport.resolveSingleCellReference(
                sheet,
                ExcelChartSourceSupport.normalizeFormula(formula.formula()),
                "Series title formula");
        yield new PreparedSeriesTitleFormula(
            ExcelChartSourceSupport.scalarText(sheet, reference), reference);
      }
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

  private static IllegalArgumentException typeChangeNotSupported(
      XSSFSheet sheet, String chartName) {
    return new IllegalArgumentException(
        "Changing chart type for existing chart '"
            + chartName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is unsupported; delete and recreate the chart to preserve unsupported detail.");
  }

  private static void applyCommonChartState(XSSFChart chart, PreparedChartDefinition definition) {
    applyChartTitle(chart, definition.title());
    applyChartLegend(chart, definition.legend());
    chart.displayBlanksAs(ExcelChartPoiBridge.toPoiDisplayBlanks(definition.displayBlanksAs()));
    chart.setPlotOnlyVisibleCells(definition.plotOnlyVisibleCells());
  }

  private static void initializeBarChart(XSSFChart chart, PreparedBarChart bar) {
    org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis categoryAxis =
        chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
    org.apache.poi.xddf.usermodel.chart.XDDFValueAxis valueAxis =
        chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
    valueAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);
    org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData)
            chart.createData(
                org.apache.poi.xddf.usermodel.chart.ChartTypes.BAR, categoryAxis, valueAxis);
    data.setBarDirection(ExcelChartPoiBridge.toPoiBarDirection(bar.barDirection()));
    data.setVaryColors(bar.varyColors());
    syncSeries(chart, data, bar.series());
  }

  private static void initializeLineChart(XSSFChart chart, PreparedLineChart line) {
    org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis categoryAxis =
        chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
    org.apache.poi.xddf.usermodel.chart.XDDFValueAxis valueAxis =
        chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
    valueAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);
    org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData)
            chart.createData(
                org.apache.poi.xddf.usermodel.chart.ChartTypes.LINE, categoryAxis, valueAxis);
    data.setVaryColors(line.varyColors());
    syncSeries(chart, data, line.series());
  }

  private static void initializePieChart(XSSFChart chart, PreparedPieChart pie) {
    org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData)
            chart.createData(org.apache.poi.xddf.usermodel.chart.ChartTypes.PIE, null, null);
    data.setVaryColors(pie.varyColors());
    data.setFirstSliceAngle(pie.firstSliceAngle());
    syncSeries(chart, data, pie.series());
  }

  private static void mutateBarChart(XSSFChart chart, PreparedBarChart bar) {
    org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData) chart.getChartSeries().getFirst();
    data.setBarDirection(ExcelChartPoiBridge.toPoiBarDirection(bar.barDirection()));
    data.setVaryColors(bar.varyColors());
    syncSeries(chart, data, bar.series());
  }

  private static void mutateLineChart(XSSFChart chart, PreparedLineChart line) {
    org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData) chart.getChartSeries().getFirst();
    data.setVaryColors(line.varyColors());
    syncSeries(chart, data, line.series());
  }

  private static void mutatePieChart(XSSFChart chart, PreparedPieChart pie) {
    org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data =
        (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData) chart.getChartSeries().getFirst();
    data.setVaryColors(pie.varyColors());
    data.setFirstSliceAngle(pie.firstSliceAngle());
    syncSeries(chart, data, pie.series());
  }

  private static void syncSeries(
      XSSFChart chart,
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data,
      List<PreparedChartSeries> definitions) {
    while (data.getSeriesCount() > definitions.size()) {
      data.removeSeries(data.getSeriesCount() - 1);
    }
    for (int index = 0; index < definitions.size(); index++) {
      PreparedChartSeries definition = definitions.get(index);
      org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series series;
      if (index < data.getSeriesCount()) {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) data.getSeries(index);
        series.replaceData(definition.categories(), definition.values());
      } else {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series)
                data.addSeries(definition.categories(), definition.values());
      }
      applySeriesTitle(series, definition.title());
    }
    reindexBarSeries(data);
    chart.plot(data);
  }

  private static void syncSeries(
      XSSFChart chart,
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data,
      List<PreparedChartSeries> definitions) {
    while (data.getSeriesCount() > definitions.size()) {
      data.removeSeries(data.getSeriesCount() - 1);
    }
    for (int index = 0; index < definitions.size(); index++) {
      PreparedChartSeries definition = definitions.get(index);
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series series;
      if (index < data.getSeriesCount()) {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) data.getSeries(index);
        series.replaceData(definition.categories(), definition.values());
      } else {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series)
                data.addSeries(definition.categories(), definition.values());
      }
      applySeriesTitle(series, definition.title());
    }
    reindexLineSeries(data);
    chart.plot(data);
  }

  private static void syncSeries(
      XSSFChart chart,
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data,
      List<PreparedChartSeries> definitions) {
    while (data.getSeriesCount() > definitions.size()) {
      data.removeSeries(data.getSeriesCount() - 1);
    }
    for (int index = 0; index < definitions.size(); index++) {
      PreparedChartSeries definition = definitions.get(index);
      org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series series;
      if (index < data.getSeriesCount()) {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) data.getSeries(index);
        series.replaceData(definition.categories(), definition.values());
      } else {
        series =
            (org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series)
                data.addSeries(definition.categories(), definition.values());
      }
      applySeriesTitle(series, definition.title());
    }
    reindexPieSeries(data);
    chart.plot(data);
  }

  private static PreparedChartTitle prepareChartTitle(
      XSSFSheet sheet, ExcelChartDefinition.Title title) {
    return switch (title) {
      case ExcelChartDefinition.Title.None _ -> new PreparedChartTitleNone();
      case ExcelChartDefinition.Title.Text text -> new PreparedChartTitleText(text.text());
      case ExcelChartDefinition.Title.Formula formula -> {
        CellReference reference =
            ExcelChartSourceSupport.resolveSingleCellReference(
                sheet,
                ExcelChartSourceSupport.normalizeFormula(formula.formula()),
                "Chart title formula");
        yield new PreparedChartTitleFormula(
            ExcelChartSourceSupport.scalarText(sheet, reference), reference);
      }
    };
  }

  private static List<PreparedChartSeries> prepareChartSeries(
      XSSFSheet sheet, List<ExcelChartDefinition.Series> definitions) {
    List<PreparedChartSeries> prepared = new ArrayList<>();
    for (ExcelChartDefinition.Series definition : definitions) {
      prepared.add(
          new PreparedChartSeries(
              prepareSeriesTitle(sheet, definition.title()),
              ExcelChartSourceSupport.toCategoryDataSource(
                  sheet, definition.categories().formula()),
              ExcelChartSourceSupport.toValueDataSource(sheet, definition.values().formula())));
    }
    return List.copyOf(prepared);
  }

  private static void applyChartTitle(XSSFChart chart, PreparedChartTitle title) {
    switch (title) {
      case PreparedChartTitleNone _ -> chart.removeTitle();
      case PreparedChartTitleText text -> chart.setTitleText(text.text());
      case PreparedChartTitleFormula formula -> applyChartTitleFormula(chart, formula);
    }
  }

  private static void applyChartTitleFormula(XSSFChart chart, PreparedChartTitleFormula formula) {
    Objects.requireNonNull(chart, "chart must not be null");
    Objects.requireNonNull(formula, "formula must not be null");

    org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle title =
        chart.getCTChart().isSetTitle()
            ? chart.getCTChart().getTitle()
            : chart.getCTChart().addNewTitle();
    org.openxmlformats.schemas.drawingml.x2006.chart.CTTx text =
        title.isSetTx() ? title.getTx() : title.addNewTx();
    if (text.isSetRich()) {
      text.unsetRich();
    }

    org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef reference =
        text.isSetStrRef() ? text.getStrRef() : text.addNewStrRef();
    reference.setF(formula.reference().formatAsString());
    writeStringReferenceCache(reference, formula.cachedText());
  }

  private static void writeStringReferenceCache(
      org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef reference, String cachedText) {
    Objects.requireNonNull(reference, "reference must not be null");
    Objects.requireNonNull(cachedText, "cachedText must not be null");

    if (reference.isSetStrCache()) {
      reference.unsetStrCache();
    }
    org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData cache = reference.addNewStrCache();
    cache.addNewPtCount().setVal(1);
    cache.addNewPt().setIdx(0);
    cache.getPtArray(0).setV(cachedText);
  }

  private static void applyChartLegend(XSSFChart chart, ExcelChartDefinition.Legend legend) {
    switch (legend) {
      case ExcelChartDefinition.Legend.Hidden _ -> chart.deleteLegend();
      case ExcelChartDefinition.Legend.Visible visible ->
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

  private static void reindexBarSeries(org.apache.poi.xddf.usermodel.chart.XDDFBarChartData data) {
    for (int index = 0; index < data.getSeriesCount(); index++) {
      ((org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) data.getSeries(index))
          .getCTBarSer()
          .getIdx()
          .setVal(index);
      ((org.apache.poi.xddf.usermodel.chart.XDDFBarChartData.Series) data.getSeries(index))
          .getCTBarSer()
          .getOrder()
          .setVal(index);
    }
  }

  private static void reindexLineSeries(
      org.apache.poi.xddf.usermodel.chart.XDDFLineChartData data) {
    for (int index = 0; index < data.getSeriesCount(); index++) {
      ((org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) data.getSeries(index))
          .getCTLineSer()
          .getIdx()
          .setVal(index);
      ((org.apache.poi.xddf.usermodel.chart.XDDFLineChartData.Series) data.getSeries(index))
          .getCTLineSer()
          .getOrder()
          .setVal(index);
    }
  }

  private static void reindexPieSeries(org.apache.poi.xddf.usermodel.chart.XDDFPieChartData data) {
    for (int index = 0; index < data.getSeriesCount(); index++) {
      ((org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) data.getSeries(index))
          .getCTPieSer()
          .getIdx()
          .setVal(index);
      ((org.apache.poi.xddf.usermodel.chart.XDDFPieChartData.Series) data.getSeries(index))
          .getCTPieSer()
          .getOrder()
          .setVal(index);
    }
  }
}
