package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectPayloadReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectReport;
import dev.erst.gridgrind.excel.ExcelChartSnapshot;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot;
import java.util.Base64;

/** Converts drawing and chart workbook snapshots into protocol report records. */
final class InspectionResultDrawingReportSupport {
  private InspectionResultDrawingReportSupport() {}

  static DrawingObjectReport toDrawingObjectReport(ExcelDrawingObjectSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelDrawingObjectSnapshot.Picture picture ->
          new DrawingObjectReport.Picture(
              picture.name(),
              toDrawingAnchorReport(picture.anchor()),
              picture.format(),
              picture.contentType(),
              picture.byteSize(),
              picture.sha256(),
              picture.widthPixels(),
              picture.heightPixels(),
              picture.description());
      case ExcelDrawingObjectSnapshot.Chart chart ->
          new DrawingObjectReport.Chart(
              chart.name(),
              toDrawingAnchorReport(chart.anchor()),
              chart.supported(),
              chart.plotTypeTokens(),
              chart.title());
      case ExcelDrawingObjectSnapshot.Shape shape ->
          new DrawingObjectReport.Shape(
              shape.name(),
              toDrawingAnchorReport(shape.anchor()),
              shape.kind(),
              shape.presetGeometryToken(),
              shape.text(),
              shape.childCount());
      case ExcelDrawingObjectSnapshot.EmbeddedObject embeddedObject ->
          new DrawingObjectReport.EmbeddedObject(
              embeddedObject.name(),
              toDrawingAnchorReport(embeddedObject.anchor()),
              embeddedObject.packagingKind(),
              embeddedObject.label(),
              embeddedObject.fileName(),
              embeddedObject.command(),
              embeddedObject.contentType(),
              embeddedObject.byteSize(),
              embeddedObject.sha256(),
              embeddedObject.previewFormat(),
              embeddedObject.previewByteSize(),
              embeddedObject.previewSha256());
      case ExcelDrawingObjectSnapshot.SignatureLine signatureLine ->
          new DrawingObjectReport.SignatureLine(
              signatureLine.name(),
              toDrawingAnchorReport(signatureLine.anchor()),
              signatureLine.setupId(),
              signatureLine.allowComments(),
              signatureLine.signingInstructions(),
              signatureLine.suggestedSigner(),
              signatureLine.suggestedSigner2(),
              signatureLine.suggestedSignerEmail(),
              signatureLine.previewFormat(),
              signatureLine.previewContentType(),
              signatureLine.previewByteSize(),
              signatureLine.previewSha256(),
              signatureLine.previewWidthPixels(),
              signatureLine.previewHeightPixels());
    };
  }

  static ChartReport toChartReport(ExcelChartSnapshot snapshot) {
    return new ChartReport(
        snapshot.name(),
        toDrawingAnchorReport(snapshot.anchor()),
        toChartTitleReport(snapshot.title()),
        toChartLegendReport(snapshot.legend()),
        snapshot.displayBlanksAs(),
        snapshot.plotOnlyVisibleCells(),
        snapshot.plots().stream()
            .map(InspectionResultDrawingReportSupport::toChartPlotReport)
            .toList());
  }

  static DrawingObjectPayloadReport toDrawingObjectPayloadReport(
      ExcelDrawingObjectPayload payload) {
    return switch (payload) {
      case ExcelDrawingObjectPayload.Picture picture ->
          new DrawingObjectPayloadReport.Picture(
              picture.name(),
              picture.format(),
              picture.contentType(),
              picture.fileName(),
              picture.sha256(),
              Base64.getEncoder().encodeToString(picture.data().bytes()),
              picture.description());
      case ExcelDrawingObjectPayload.EmbeddedObject embeddedObject ->
          new DrawingObjectPayloadReport.EmbeddedObject(
              embeddedObject.name(),
              embeddedObject.packagingKind(),
              embeddedObject.contentType(),
              embeddedObject.fileName(),
              embeddedObject.sha256(),
              Base64.getEncoder().encodeToString(embeddedObject.data().bytes()),
              embeddedObject.label(),
              embeddedObject.command());
    };
  }

  static DrawingAnchorReport toDrawingAnchorReport(ExcelDrawingAnchor anchor) {
    return switch (anchor) {
      case ExcelDrawingAnchor.TwoCell twoCell ->
          new DrawingAnchorReport.TwoCell(
              toDrawingMarkerReport(twoCell.from()),
              toDrawingMarkerReport(twoCell.to()),
              twoCell.behavior());
      case ExcelDrawingAnchor.OneCell oneCell ->
          new DrawingAnchorReport.OneCell(
              toDrawingMarkerReport(oneCell.from()),
              oneCell.widthEmu(),
              oneCell.heightEmu(),
              oneCell.behavior());
      case ExcelDrawingAnchor.Absolute absolute ->
          new DrawingAnchorReport.Absolute(
              absolute.xEmu(),
              absolute.yEmu(),
              absolute.widthEmu(),
              absolute.heightEmu(),
              absolute.behavior());
    };
  }

  static DrawingMarkerReport toDrawingMarkerReport(ExcelDrawingMarker marker) {
    return new DrawingMarkerReport(
        marker.columnIndex(), marker.rowIndex(), marker.dx(), marker.dy());
  }

  private static ChartReport.Title toChartTitleReport(ExcelChartSnapshot.Title title) {
    return switch (title) {
      case ExcelChartSnapshot.Title.None _ -> new ChartReport.Title.None();
      case ExcelChartSnapshot.Title.Text text -> new ChartReport.Title.Text(text.text());
      case ExcelChartSnapshot.Title.Formula formula ->
          new ChartReport.Title.Formula(formula.formula(), formula.cachedText());
    };
  }

  private static ChartReport.Legend toChartLegendReport(ExcelChartSnapshot.Legend legend) {
    return switch (legend) {
      case ExcelChartSnapshot.Legend.Hidden _ -> new ChartReport.Legend.Hidden();
      case ExcelChartSnapshot.Legend.Visible visible ->
          new ChartReport.Legend.Visible(visible.position());
    };
  }

  private static ChartReport.Axis toChartAxisReport(ExcelChartSnapshot.Axis axis) {
    return new ChartReport.Axis(axis.kind(), axis.position(), axis.crosses(), axis.visible());
  }

  private static ChartReport.Series toChartSeriesReport(ExcelChartSnapshot.Series series) {
    return new ChartReport.Series(
        toChartTitleReport(series.title()),
        toChartDataSourceReport(series.categories()),
        toChartDataSourceReport(series.values()),
        series.smooth(),
        series.markerStyle(),
        series.markerSize(),
        series.explosion());
  }

  private static ChartReport.DataSource toChartDataSourceReport(
      ExcelChartSnapshot.DataSource source) {
    return switch (source) {
      case ExcelChartSnapshot.DataSource.StringReference reference ->
          new ChartReport.DataSource.StringReference(reference.formula(), reference.cachedValues());
      case ExcelChartSnapshot.DataSource.NumericReference reference ->
          new ChartReport.DataSource.NumericReference(
              reference.formula(), reference.formatCode(), reference.cachedValues());
      case ExcelChartSnapshot.DataSource.StringLiteral literal ->
          new ChartReport.DataSource.StringLiteral(literal.values());
      case ExcelChartSnapshot.DataSource.NumericLiteral literal ->
          new ChartReport.DataSource.NumericLiteral(literal.formatCode(), literal.values());
    };
  }

  private static ChartReport.Plot toChartPlotReport(ExcelChartSnapshot.Plot plot) {
    return switch (plot) {
      case ExcelChartSnapshot.Area area ->
          new ChartReport.Area(
              area.varyColors(),
              area.grouping(),
              area.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              area.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Area3D area3D ->
          new ChartReport.Area3D(
              area3D.varyColors(),
              area3D.grouping(),
              area3D.gapDepth(),
              area3D.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              area3D.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Bar bar ->
          new ChartReport.Bar(
              bar.varyColors(),
              bar.barDirection(),
              bar.grouping(),
              bar.gapWidth(),
              bar.overlap(),
              bar.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              bar.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Bar3D bar3D ->
          new ChartReport.Bar3D(
              bar3D.varyColors(),
              bar3D.barDirection(),
              bar3D.grouping(),
              bar3D.gapDepth(),
              bar3D.gapWidth(),
              bar3D.shape(),
              bar3D.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              bar3D.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Doughnut doughnut ->
          new ChartReport.Doughnut(
              doughnut.varyColors(),
              doughnut.firstSliceAngle(),
              doughnut.holeSize(),
              doughnut.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Line line ->
          new ChartReport.Line(
              line.varyColors(),
              line.grouping(),
              line.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              line.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Line3D line3D ->
          new ChartReport.Line3D(
              line3D.varyColors(),
              line3D.grouping(),
              line3D.gapDepth(),
              line3D.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              line3D.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Pie pie ->
          new ChartReport.Pie(
              pie.varyColors(),
              pie.firstSliceAngle(),
              pie.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Pie3D pie3D ->
          new ChartReport.Pie3D(
              pie3D.varyColors(),
              pie3D.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Radar radar ->
          new ChartReport.Radar(
              radar.varyColors(),
              radar.style(),
              radar.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              radar.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Scatter scatter ->
          new ChartReport.Scatter(
              scatter.varyColors(),
              scatter.style(),
              scatter.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              scatter.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Surface surface ->
          new ChartReport.Surface(
              surface.varyColors(),
              surface.wireframe(),
              surface.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              surface.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Surface3D surface3D ->
          new ChartReport.Surface3D(
              surface3D.varyColors(),
              surface3D.wireframe(),
              surface3D.axes().stream()
                  .map(InspectionResultDrawingReportSupport::toChartAxisReport)
                  .toList(),
              surface3D.series().stream()
                  .map(InspectionResultDrawingReportSupport::toChartSeriesReport)
                  .toList());
      case ExcelChartSnapshot.Unsupported unsupported ->
          new ChartReport.Unsupported(unsupported.plotTypeToken(), unsupported.detail());
    };
  }
}
