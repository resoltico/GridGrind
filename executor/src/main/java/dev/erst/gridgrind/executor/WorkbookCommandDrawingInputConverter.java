package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ChartAxisInput;
import dev.erst.gridgrind.contract.dto.ChartDataSourceInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartLegendInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.excel.ExcelBinaryData;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelCustomXmlImportDefinition;
import dev.erst.gridgrind.excel.ExcelCustomXmlMappingLocator;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectDefinition;
import dev.erst.gridgrind.excel.ExcelPictureDefinition;
import dev.erst.gridgrind.excel.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.ExcelSignatureLineDefinition;

/** Converts drawing-oriented structured contract inputs. */
final class WorkbookCommandDrawingInputConverter {
  private WorkbookCommandDrawingInputConverter() {}

  static ExcelCustomXmlImportDefinition toExcelCustomXmlImportDefinition(
      CustomXmlImportInput input) {
    return new ExcelCustomXmlImportDefinition(
        toExcelCustomXmlMappingLocator(input.locator()),
        WorkbookCommandSourceSupport.inlineText(input.xml(), "custom XML"));
  }

  static ExcelCustomXmlMappingLocator toExcelCustomXmlMappingLocator(
      dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator locator) {
    return new ExcelCustomXmlMappingLocator(locator.mapId(), locator.name());
  }

  static ExcelPictureDefinition toExcelPictureDefinition(PictureInput picture) {
    return new ExcelPictureDefinition(
        picture.name(),
        new ExcelBinaryData(
            WorkbookCommandSourceSupport.inlineBinary(picture.image().source(), "picture payload")),
        picture.image().format(),
        toExcelDrawingAnchor(picture.anchor()),
        picture.description() == null
            ? null
            : WorkbookCommandSourceSupport.inlineText(
                picture.description(), "picture description"));
  }

  static ExcelChartDefinition toExcelChartDefinition(ChartInput chart) {
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
      SignatureLineInput signatureLine) {
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
        signatureLine.plainSignature().isEmpty()
            ? null
            : signatureLine.plainSignature().orElseThrow().format(),
        signatureLine.plainSignature().isEmpty()
            ? null
            : new ExcelBinaryData(
                WorkbookCommandSourceSupport.inlineBinary(
                    signatureLine.plainSignature().orElseThrow().source(),
                    "signature-line plain signature")));
  }

  static ExcelShapeDefinition toExcelShapeDefinition(ShapeInput shape) {
    return new ExcelShapeDefinition(
        shape.name(),
        shape.kind(),
        toExcelDrawingAnchor(shape.anchor()),
        shape.presetGeometryToken(),
        shape.text() == null
            ? null
            : WorkbookCommandSourceSupport.inlineText(shape.text(), "shape text"));
  }

  static ExcelEmbeddedObjectDefinition toExcelEmbeddedObjectDefinition(
      EmbeddedObjectInput embeddedObject) {
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

  static ExcelDrawingAnchor.TwoCell toExcelDrawingAnchor(DrawingAnchorInput anchor) {
    return switch (anchor) {
      case DrawingAnchorInput.TwoCell twoCell ->
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

  private static ExcelChartDefinition.Title toExcelChartTitle(ChartTitleInput title) {
    return switch (title) {
      case ChartTitleInput.None _ -> new ExcelChartDefinition.Title.None();
      case ChartTitleInput.Text text ->
          new ExcelChartDefinition.Title.Text(
              WorkbookCommandSourceSupport.inlineText(text.source(), "chart title"));
      case ChartTitleInput.Formula formula ->
          new ExcelChartDefinition.Title.Formula(formula.formula());
    };
  }

  private static ExcelChartDefinition.Legend toExcelChartLegend(ChartLegendInput legend) {
    return switch (legend) {
      case ChartLegendInput.Hidden _ -> new ExcelChartDefinition.Legend.Hidden();
      case ChartLegendInput.Visible visible ->
          new ExcelChartDefinition.Legend.Visible(visible.position());
    };
  }

  private static ExcelChartDefinition.Series toExcelChartSeries(ChartSeriesInput series) {
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
      ChartDataSourceInput source) {
    return switch (source) {
      case ChartDataSourceInput.Reference reference ->
          new ExcelChartDefinition.DataSource.Reference(reference.formula());
      case ChartDataSourceInput.StringLiteral literal ->
          new ExcelChartDefinition.DataSource.StringLiteral(literal.values());
      case ChartDataSourceInput.NumericLiteral literal ->
          new ExcelChartDefinition.DataSource.NumericLiteral(literal.values());
    };
  }

  private static ExcelChartDefinition.Axis toExcelChartAxis(ChartAxisInput axis) {
    return new ExcelChartDefinition.Axis(
        axis.kind(), axis.position(), axis.crosses(), axis.visible());
  }

  private static ExcelChartDefinition.Plot toExcelChartPlot(ChartPlotInput plot) {
    return switch (plot) {
      case ChartPlotInput.Area area ->
          new ExcelChartDefinition.Area(
              area.varyColors(),
              area.grouping(),
              area.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              area.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotInput.Area3D area3D ->
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
      case ChartPlotInput.Bar bar ->
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
      case ChartPlotInput.Bar3D bar3D ->
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
      case ChartPlotInput.Doughnut doughnut ->
          new ExcelChartDefinition.Doughnut(
              doughnut.varyColors(),
              doughnut.firstSliceAngle(),
              doughnut.holeSize(),
              doughnut.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotInput.Line line ->
          new ExcelChartDefinition.Line(
              line.varyColors(),
              line.grouping(),
              line.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              line.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotInput.Line3D line3D ->
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
      case ChartPlotInput.Pie pie ->
          new ExcelChartDefinition.Pie(
              pie.varyColors(),
              pie.firstSliceAngle(),
              pie.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotInput.Pie3D pie3D ->
          new ExcelChartDefinition.Pie3D(
              pie3D.varyColors(),
              pie3D.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotInput.Radar radar ->
          new ExcelChartDefinition.Radar(
              radar.varyColors(),
              radar.style(),
              radar.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              radar.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotInput.Scatter scatter ->
          new ExcelChartDefinition.Scatter(
              scatter.varyColors(),
              scatter.style(),
              scatter.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              scatter.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotInput.Surface surface ->
          new ExcelChartDefinition.Surface(
              surface.varyColors(),
              surface.wireframe(),
              surface.axes().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartAxis)
                  .toList(),
              surface.series().stream()
                  .map(WorkbookCommandDrawingInputConverter::toExcelChartSeries)
                  .toList());
      case ChartPlotInput.Surface3D surface3D ->
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
