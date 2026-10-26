package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.chart.ChartAxisSpec;
import dev.erst.gridgrind.core.model.chart.ChartDataSourceSpec;
import dev.erst.gridgrind.core.model.chart.ChartLegendSpec;
import dev.erst.gridgrind.core.model.chart.ChartPlotSpec;
import dev.erst.gridgrind.core.model.chart.ChartSeriesSpec;
import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.customxml.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.drawing.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.drawing.ExcelSignatureLineDefinition;

/** Converts drawing-oriented structured contract inputs. */
final class WorkbookCommandDrawingInputConverter {
  private WorkbookCommandDrawingInputConverter() {}

  static ExcelCustomXmlImportDefinition toExcelCustomXmlImportDefinition(
      CustomXmlImportSpec input) {
    return new ExcelCustomXmlImportDefinition(
        input.locator(), WorkbookCommandSourceSupport.inlineText(input.xml(), "custom XML"));
  }

  static ExcelPictureDefinition toExcelPictureDefinition(PictureSpec picture) {
    return new ExcelPictureDefinition(
        picture.name(),
        new ExcelBinaryData(
            WorkbookCommandSourceSupport.inlineBinary(picture.image().source(), "picture payload")),
        picture.image().format(),
        toExcelDrawingAnchor(picture.anchor()),
        picture
            .description()
            .map(
                description ->
                    WorkbookCommandSourceSupport.inlineText(description, "picture description")));
  }

  static ExcelChartDefinition toExcelChartDefinition(ChartSpec chart) {
    return new ExcelChartDefinition(
        chart.name(),
        toExcelDrawingAnchor(chart.anchor()),
        toExcelChartTitle(chart.title()),
        toExcelChartLegend(chart.legend()),
        chart.displayBlanksAs(),
        chart.plotOnlyVisibleCells(),
        chart.plots().stream()
            .map(WorkbookCommandDrawingInputConverter::toExcelChartPlot)
            .toList());
  }

  static ExcelSignatureLineDefinition toExcelSignatureLineDefinition(
      SignatureLineSpec signatureLine) {
    return new ExcelSignatureLineDefinition(
        signatureLine.name(),
        toExcelDrawingAnchor(signatureLine.anchor()),
        signatureLine.allowComments(),
        signatureLine.signingInstructions().orElse(null),
        signatureLine.suggestedSigner().orElse(null),
        signatureLine.suggestedSigner2().orElse(null),
        signatureLine.suggestedSignerEmail().orElse(null),
        signatureLine.caption().orElse(null),
        signatureLine.invalidStamp().orElse(null),
        signatureLine.plainSignature().map(PictureDataSpec::format),
        signatureLine
            .plainSignature()
            .map(
                plainSignature ->
                    new ExcelBinaryData(
                        WorkbookCommandSourceSupport.inlineBinary(
                            plainSignature.source(), "signature-line plain signature"))));
  }

  static ExcelShapeDefinition toExcelShapeDefinition(ShapeSpec shape) {
    return switch (shape) {
      case ShapeSpec.SimpleShape simpleShape ->
          new ExcelShapeDefinition.SimpleShape(
              simpleShape.name(),
              toExcelDrawingAnchor(simpleShape.anchor()),
              simpleShape.presetGeometryToken(),
              simpleShape
                  .text()
                  .map(text -> WorkbookCommandSourceSupport.inlineText(text, "shape text")));
      case ShapeSpec.Connector connector ->
          new ExcelShapeDefinition.Connector(
              connector.name(), toExcelDrawingAnchor(connector.anchor()));
    };
  }

  static ExcelEmbeddedObjectDefinition toExcelEmbeddedObjectDefinition(
      EmbeddedObjectSpec embeddedObject) {
    return new ExcelEmbeddedObjectDefinition(
        embeddedObject.name(),
        embeddedObject.label(),
        embeddedObject.fileName(),
        embeddedObject.command(),
        new ExcelBinaryData(
            WorkbookCommandSourceSupport.inlineBinary(
                embeddedObject.payload(), "embedded-object payload")),
        embeddedObject.previewImage().format(),
        new ExcelBinaryData(
            WorkbookCommandSourceSupport.inlineBinary(
                embeddedObject.previewImage().source(), "embedded-object preview image")),
        toExcelDrawingAnchor(embeddedObject.anchor()));
  }

  static ExcelDrawingAnchor.TwoCell toExcelDrawingAnchor(DrawingAnchorSpec anchor) {
    return switch (anchor) {
      case DrawingAnchorSpec.TwoCell twoCell ->
          new ExcelDrawingAnchor.TwoCell(
              new ExcelDrawingMarker(
                  twoCell.from().columnIndex(),
                  twoCell.from().rowIndex(),
                  twoCell.from().dx(),
                  twoCell.from().dy()),
              new ExcelDrawingMarker(
                  twoCell.to().columnIndex(),
                  twoCell.to().rowIndex(),
                  twoCell.to().dx(),
                  twoCell.to().dy()),
              twoCell.behavior());
    };
  }

  private static ExcelChartDefinition.Title toExcelChartTitle(ChartTitleSpec title) {
    return switch (title) {
      case ChartTitleSpec.None _ -> new ExcelChartDefinition.Title.None();
      case ChartTitleSpec.Text text ->
          new ExcelChartDefinition.Title.Text(
              WorkbookCommandSourceSupport.inlineText(text.source(), "chart title"));
      case ChartTitleSpec.Formula formula ->
          new ExcelChartDefinition.Title.Formula(formula.formula());
    };
  }

  private static ExcelChartDefinition.Legend toExcelChartLegend(ChartLegendSpec legend) {
    return switch (legend) {
      case ChartLegendSpec.Hidden _ -> new ExcelChartDefinition.Legend.Hidden();
      case ChartLegendSpec.Visible visible ->
          new ExcelChartDefinition.Legend.Visible(visible.position());
    };
  }

  private static ExcelChartDefinition.Series toExcelChartSeries(ChartSeriesSpec series) {
    return new ExcelChartDefinition.Series(
        toExcelChartTitle(series.title()),
        toExcelChartDataSource(series.categories()),
        toExcelChartDataSource(series.values()),
        series.smooth(),
        series.markerStyle(),
        series.markerSize(),
        series.explosion());
  }

  private static ExcelChartDefinition.DataSource toExcelChartDataSource(
      ChartDataSourceSpec source) {
    return switch (source) {
      case ChartDataSourceSpec.Reference reference ->
          new ExcelChartDefinition.DataSource.Reference(reference.formula());
      case ChartDataSourceSpec.StringLiteral literal ->
          new ExcelChartDefinition.DataSource.StringLiteral(literal.values());
      case ChartDataSourceSpec.NumericLiteral literal ->
          new ExcelChartDefinition.DataSource.NumericLiteral(literal.values());
    };
  }

  private static ExcelChartDefinition.Axis toExcelChartAxis(ChartAxisSpec axis) {
    return new ExcelChartDefinition.Axis(
        axis.kind(), axis.position(), axis.crosses(), axis.visible());
  }

  private static ExcelChartDefinition.Plot toExcelChartPlot(ChartPlotSpec plot) {
    return switch (plot) {
      case ChartPlotSpec.Area area ->
          new ExcelChartDefinition.Area(
              area.varyColors(),
              area.grouping(),
              area.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              area.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Area3D area3D ->
          new ExcelChartDefinition.Area3D(
              area3D.varyColors(),
              area3D.grouping(),
              area3D.gapDepth(),
              area3D.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              area3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Bar bar ->
          new ExcelChartDefinition.Bar(
              bar.varyColors(),
              bar.barDirection(),
              bar.grouping(),
              bar.gapWidth(),
              bar.overlap(),
              bar.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              bar.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Bar3D bar3D ->
          new ExcelChartDefinition.Bar3D(
              bar3D.varyColors(),
              bar3D.barDirection(),
              bar3D.grouping(),
              bar3D.gapDepth(),
              bar3D.gapWidth(),
              bar3D.shape(),
              bar3D.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              bar3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Doughnut doughnut ->
          new ExcelChartDefinition.Doughnut(
              doughnut.varyColors(),
              doughnut.firstSliceAngle(),
              doughnut.holeSize(),
              doughnut.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Line line ->
          new ExcelChartDefinition.Line(
              line.varyColors(),
              line.grouping(),
              line.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              line.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Line3D line3D ->
          new ExcelChartDefinition.Line3D(
              line3D.varyColors(),
              line3D.grouping(),
              line3D.gapDepth(),
              line3D.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              line3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Pie pie ->
          new ExcelChartDefinition.Pie(
              pie.varyColors(),
              pie.firstSliceAngle(),
              pie.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Pie3D pie3D ->
          new ExcelChartDefinition.Pie3D(
              pie3D.varyColors(),
              pie3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Radar radar ->
          new ExcelChartDefinition.Radar(
              radar.varyColors(),
              radar.style(),
              radar.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              radar.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Scatter scatter ->
          new ExcelChartDefinition.Scatter(
              scatter.varyColors(),
              scatter.style(),
              scatter.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              scatter.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Surface surface ->
          new ExcelChartDefinition.Surface(
              surface.varyColors(),
              surface.wireframe(),
              surface.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              surface.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotSpec.Surface3D surface3D ->
          new ExcelChartDefinition.Surface3D(
              surface3D.varyColors(),
              surface3D.wireframe(),
              surface3D.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              surface3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
    };
  }
}
