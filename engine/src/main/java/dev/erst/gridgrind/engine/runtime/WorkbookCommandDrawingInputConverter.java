package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.chart.ResolvedChart;
import dev.erst.gridgrind.core.model.chart.ResolvedChartTitle;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.customxml.ResolvedCustomXmlImport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.source.BinaryContent;

/** Converts drawing-oriented structured contract inputs. */
final class WorkbookCommandDrawingInputConverter {
  private WorkbookCommandDrawingInputConverter() {}

  static ResolvedCustomXmlImport toResolvedCustomXmlImport(CustomXmlImportSpec input) {
    return new ResolvedCustomXmlImport(
        input.locator(), WorkbookCommandSourceSupport.inlineText(input.xml(), "custom XML"));
  }

  static ResolvedPicture toResolvedPicture(PictureSpec picture) {
    return new ResolvedPicture(
        picture.name(),
        new BinaryContent(
            WorkbookCommandSourceSupport.inlineBinary(picture.image().source(), "picture payload")),
        picture.image().format(),
        twoCellAnchor(picture.anchor()),
        picture
            .description()
            .map(
                description ->
                    WorkbookCommandSourceSupport.inlineText(description, "picture description")));
  }

  static ResolvedChart toResolvedChart(ChartSpec chart) {
    return new ResolvedChart(
        chart.name(),
        twoCellAnchor(chart.anchor()),
        toExcelChartTitle(chart.title()),
        chart.legend(),
        chart.displayBlanksAs(),
        chart.plotOnlyVisibleCells(),
        chart.plots().stream()
            .map(WorkbookCommandDrawingInputConverter::toExcelChartPlot)
            .toList());
  }

  static ResolvedSignatureLine toResolvedSignatureLine(SignatureLineSpec signatureLine) {
    return new ResolvedSignatureLine(
        signatureLine.name(),
        twoCellAnchor(signatureLine.anchor()),
        signatureLine.allowComments(),
        signatureLine.signingInstructions(),
        signatureLine.suggestedSigner(),
        signatureLine.suggestedSigner2(),
        signatureLine.suggestedSignerEmail(),
        signatureLine.caption(),
        signatureLine.invalidStamp(),
        signatureLine.plainSignature().map(PictureDataSpec::format),
        signatureLine
            .plainSignature()
            .map(
                plainSignature ->
                    new BinaryContent(
                        WorkbookCommandSourceSupport.inlineBinary(
                            plainSignature.source(), "signature-line plain signature"))));
  }

  static ResolvedShape toResolvedShape(ShapeSpec shape) {
    return switch (shape) {
      case ShapeSpec.SimpleShape simpleShape ->
          new ResolvedShape.SimpleShape(
              simpleShape.name(),
              twoCellAnchor(simpleShape.anchor()),
              simpleShape.presetGeometryToken(),
              simpleShape
                  .text()
                  .map(text -> WorkbookCommandSourceSupport.inlineText(text, "shape text")));
      case ShapeSpec.Connector connector -> connector;
    };
  }

  static ResolvedEmbeddedObject toResolvedEmbeddedObject(EmbeddedObjectSpec embeddedObject) {
    return new ResolvedEmbeddedObject(
        embeddedObject.name(),
        embeddedObject.label(),
        embeddedObject.fileName(),
        embeddedObject.command(),
        new BinaryContent(
            WorkbookCommandSourceSupport.inlineBinary(
                embeddedObject.payload(), "embedded-object payload")),
        embeddedObject.previewImage().format(),
        new BinaryContent(
            WorkbookCommandSourceSupport.inlineBinary(
                embeddedObject.previewImage().source(), "embedded-object preview image")),
        twoCellAnchor(embeddedObject.anchor()));
  }

  static DrawingAnchorSpec.TwoCell twoCellAnchor(DrawingAnchorSpec anchor) {
    return switch (anchor) {
      case DrawingAnchorSpec.TwoCell twoCell -> twoCell;
    };
  }

  private static ResolvedChartTitle toExcelChartTitle(ChartTitleSpec title) {
    return switch (title) {
      case ChartTitleSpec.None none -> none;
      case ChartTitleSpec.Text text ->
          new ResolvedChartTitle.Text(
              WorkbookCommandSourceSupport.inlineText(text.source(), "chart title"));
      case ChartTitleSpec.Formula formula -> formula;
    };
  }

  private static ResolvedChart.Series toExcelChartSeries(ChartSeriesSpec series) {
    return new ResolvedChart.Series(
        toExcelChartTitle(series.title()),
        series.categories(),
        series.values(),
        series.smooth(),
        series.markerStyle(),
        series.markerSize(),
        series.explosion());
  }

  private static ResolvedChart.Plot toExcelChartPlot(ChartPlotSpec plot) {
    return switch (plot) {
      case ChartPlotSpec.Area area ->
          new ResolvedChart.Area(
              area.varyColors(),
              area.grouping(),
              area.axes(),
              area.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Area3D area3D ->
          new ResolvedChart.Area3D(
              area3D.varyColors(),
              area3D.grouping(),
              area3D.gapDepth(),
              area3D.axes(),
              area3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Bar bar ->
          new ResolvedChart.Bar(
              bar.varyColors(),
              bar.barDirection(),
              bar.grouping(),
              bar.gapWidth(),
              bar.overlap(),
              bar.axes(),
              bar.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Bar3D bar3D ->
          new ResolvedChart.Bar3D(
              bar3D.varyColors(),
              bar3D.barDirection(),
              bar3D.grouping(),
              bar3D.gapDepth(),
              bar3D.gapWidth(),
              bar3D.shape(),
              bar3D.axes(),
              bar3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Doughnut doughnut ->
          new ResolvedChart.Doughnut(
              doughnut.varyColors(),
              doughnut.firstSliceAngle(),
              doughnut.holeSize(),
              doughnut.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Line line ->
          new ResolvedChart.Line(
              line.varyColors(),
              line.grouping(),
              line.axes(),
              line.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Line3D line3D ->
          new ResolvedChart.Line3D(
              line3D.varyColors(),
              line3D.grouping(),
              line3D.gapDepth(),
              line3D.axes(),
              line3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Pie pie ->
          new ResolvedChart.Pie(
              pie.varyColors(),
              pie.firstSliceAngle(),
              pie.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Pie3D pie3D ->
          new ResolvedChart.Pie3D(
              pie3D.varyColors(),
              pie3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Radar radar ->
          new ResolvedChart.Radar(
              radar.varyColors(),
              radar.style(),
              radar.axes(),
              radar.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Scatter scatter ->
          new ResolvedChart.Scatter(
              scatter.varyColors(),
              scatter.style(),
              scatter.axes(),
              scatter.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Surface surface ->
          new ResolvedChart.Surface(
              surface.varyColors(),
              surface.wireframe(),
              surface.axes(),
              surface.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Surface3D surface3D ->
          new ResolvedChart.Surface3D(
              surface3D.varyColors(),
              surface3D.wireframe(),
              surface3D.axes(),
              surface3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
    };
  }
}
