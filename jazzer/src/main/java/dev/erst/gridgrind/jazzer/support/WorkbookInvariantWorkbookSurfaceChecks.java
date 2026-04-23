package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectPayloadReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import java.util.HashSet;

/**
 * Owns workbook-surface invariant checks for summaries, drawing objects, charts, and pivot tables.
 */
final class WorkbookInvariantWorkbookSurfaceChecks {
  private WorkbookInvariantWorkbookSurfaceChecks() {}

  static void requireWorkbookSummaryShape(GridGrindResponse.WorkbookSummary workbook) {
    WorkbookInvariantChecks.require(workbook != null, "workbook summary must not be null");
    WorkbookInvariantChecks.require(workbook.sheetCount() >= 0, "sheetCount must not be negative");
    WorkbookInvariantChecks.require(
        workbook.namedRangeCount() >= 0, "namedRangeCount must not be negative");
    WorkbookInvariantChecks.require(workbook.sheetNames() != null, "sheetNames must not be null");
    WorkbookInvariantChecks.require(
        workbook.sheetCount() == workbook.sheetNames().size(),
        "sheetCount must match sheetNames size");
    workbook
        .sheetNames()
        .forEach(
            sheetName ->
                WorkbookInvariantChecks.require(
                    sheetName != null && !sheetName.isBlank(), "sheetName must not be blank"));
    switch (workbook) {
      case GridGrindResponse.WorkbookSummary.Empty empty -> {
        WorkbookInvariantChecks.require(
            empty.sheetCount() == 0, "empty workbook summary must have sheetCount 0");
        WorkbookInvariantChecks.require(
            empty.sheetNames().isEmpty(), "empty workbook summary must have no sheet names");
      }
      case GridGrindResponse.WorkbookSummary.WithSheets withSheets -> {
        WorkbookInvariantChecks.require(
            withSheets.sheetCount() > 0,
            "non-empty workbook summary must have positive sheetCount");
        WorkbookInvariantChecks.requireNonBlank(withSheets.activeSheetName(), "activeSheetName");
        WorkbookInvariantChecks.require(
            withSheets.sheetNames().contains(withSheets.activeSheetName()),
            "activeSheetName must be present in sheetNames");
        WorkbookInvariantChecks.require(
            withSheets.selectedSheetNames() != null, "selectedSheetNames must not be null");
        WorkbookInvariantChecks.require(
            !withSheets.selectedSheetNames().isEmpty(), "selectedSheetNames must not be empty");
        WorkbookInvariantChecks.require(
            withSheets.selectedSheetNames().size()
                == new HashSet<>(withSheets.selectedSheetNames()).size(),
            "selectedSheetNames must be unique");
        withSheets
            .selectedSheetNames()
            .forEach(
                selectedSheetName -> {
                  WorkbookInvariantChecks.requireNonBlank(selectedSheetName, "selectedSheetName");
                  WorkbookInvariantChecks.require(
                      withSheets.sheetNames().contains(selectedSheetName),
                      "selectedSheetNames must be present in sheetNames");
                });
      }
    }
  }

  static void requireSheetSummaryShape(GridGrindResponse.SheetSummaryReport sheet) {
    WorkbookInvariantChecks.require(sheet.sheetName() != null, "sheetName must not be null");
    WorkbookInvariantChecks.require(!sheet.sheetName().isBlank(), "sheetName must not be blank");
    WorkbookInvariantChecks.require(sheet.visibility() != null, "visibility must not be null");
    WorkbookInvariantChecks.require(sheet.protection() != null, "protection must not be null");
    WorkbookInvariantChecks.require(
        sheet.physicalRowCount() >= 0, "physicalRowCount must not be negative");
    WorkbookInvariantChecks.require(
        sheet.lastRowIndex() >= -1, "lastRowIndex must be greater than or equal to -1");
    WorkbookInvariantChecks.require(
        sheet.lastColumnIndex() >= -1, "lastColumnIndex must be greater than or equal to -1");
    switch (sheet.protection()) {
      case GridGrindResponse.SheetProtectionReport.Unprotected _ -> {}
      case GridGrindResponse.SheetProtectionReport.Protected protectedReport ->
          WorkbookInvariantChecks.require(
              protectedReport.settings() != null, "protected sheet settings must not be null");
    }
  }

  static void requireDrawingObjectShape(DrawingObjectReport drawingObject) {
    WorkbookInvariantChecks.require(drawingObject != null, "drawing object must not be null");
    WorkbookInvariantChecks.requireNonBlank(drawingObject.name(), "drawing object name");
    requireDrawingAnchorShape(drawingObject.anchor());
    switch (drawingObject) {
      case DrawingObjectReport.Picture picture -> requirePictureDrawingObjectShape(picture);
      case DrawingObjectReport.Chart chart -> requireChartDrawingObjectShape(chart);
      case DrawingObjectReport.Shape shape -> requireShapeDrawingObjectShape(shape);
      case DrawingObjectReport.EmbeddedObject embeddedObject ->
          requireEmbeddedDrawingObjectShape(embeddedObject);
      case DrawingObjectReport.SignatureLine signatureLine ->
          requireSignatureLineDrawingObjectShape(signatureLine);
    }
  }

  static void requirePictureDrawingObjectShape(DrawingObjectReport.Picture picture) {
    WorkbookInvariantChecks.requireNonBlank(picture.contentType(), "picture contentType");
    WorkbookInvariantChecks.requireNonBlank(picture.sha256(), "picture sha256");
    WorkbookInvariantChecks.require(
        picture.byteSize() >= 0L, "picture byteSize must not be negative");
    if (picture.widthPixels() != null) {
      WorkbookInvariantChecks.require(
          picture.widthPixels() >= 0, "picture widthPixels must not be negative");
    }
    if (picture.heightPixels() != null) {
      WorkbookInvariantChecks.require(
          picture.heightPixels() >= 0, "picture heightPixels must not be negative");
    }
    if (picture.description() != null) {
      WorkbookInvariantChecks.require(
          !picture.description().isBlank(), "picture description must not be blank");
    }
  }

  static void requireChartDrawingObjectShape(DrawingObjectReport.Chart chart) {
    WorkbookInvariantChecks.require(
        chart.plotTypeTokens() != null, "chart plotTypeTokens must not be null");
    chart
        .plotTypeTokens()
        .forEach(token -> WorkbookInvariantChecks.requireNonBlank(token, "chart plotTypeToken"));
    WorkbookInvariantChecks.require(chart.title() != null, "chart title must not be null");
  }

  static void requireShapeDrawingObjectShape(DrawingObjectReport.Shape shape) {
    if (shape.presetGeometryToken() != null) {
      WorkbookInvariantChecks.require(
          !shape.presetGeometryToken().isBlank(), "shape presetGeometryToken must not be blank");
    }
    if (shape.text() != null) {
      WorkbookInvariantChecks.require(!shape.text().isBlank(), "shape text must not be blank");
    }
    WorkbookInvariantChecks.require(
        shape.childCount() >= 0, "shape childCount must not be negative");
  }

  static void requireEmbeddedDrawingObjectShape(DrawingObjectReport.EmbeddedObject embeddedObject) {
    WorkbookInvariantChecks.requireNonBlank(
        embeddedObject.contentType(), "embedded object contentType");
    WorkbookInvariantChecks.requireNonBlank(embeddedObject.sha256(), "embedded object sha256");
    WorkbookInvariantChecks.require(
        embeddedObject.byteSize() >= 0L, "embedded object byteSize must not be negative");
    if (embeddedObject.label() != null) {
      WorkbookInvariantChecks.require(
          !embeddedObject.label().isBlank(), "embedded object label must not be blank");
    }
    if (embeddedObject.fileName() != null) {
      WorkbookInvariantChecks.require(
          !embeddedObject.fileName().isBlank(), "embedded object fileName must not be blank");
    }
    if (embeddedObject.command() != null) {
      WorkbookInvariantChecks.require(
          !embeddedObject.command().isBlank(), "embedded object command must not be blank");
    }
    if (embeddedObject.previewByteSize() != null) {
      WorkbookInvariantChecks.require(
          embeddedObject.previewByteSize() >= 0L,
          "embedded object previewByteSize must not be negative");
      WorkbookInvariantChecks.require(
          embeddedObject.previewFormat() != null,
          "embedded object previewByteSize requires previewFormat");
    }
    if (embeddedObject.previewSha256() != null) {
      WorkbookInvariantChecks.require(
          !embeddedObject.previewSha256().isBlank(),
          "embedded object previewSha256 must not be blank");
      WorkbookInvariantChecks.require(
          embeddedObject.previewFormat() != null,
          "embedded object previewSha256 requires previewFormat");
    }
  }

  static void requireSignatureLineDrawingObjectShape(
      DrawingObjectReport.SignatureLine signatureLine) {
    if (signatureLine.setupId() != null) {
      WorkbookInvariantChecks.requireNonBlank(signatureLine.setupId(), "signature line setupId");
    }
    if (signatureLine.signingInstructions() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.signingInstructions(), "signature line signingInstructions");
    }
    if (signatureLine.suggestedSigner() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.suggestedSigner(), "signature line suggestedSigner");
    }
    if (signatureLine.suggestedSigner2() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.suggestedSigner2(), "signature line suggestedSigner2");
    }
    if (signatureLine.suggestedSignerEmail() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.suggestedSignerEmail(), "signature line suggestedSignerEmail");
    }
    if (signatureLine.previewContentType() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.previewContentType(), "signature line previewContentType");
      WorkbookInvariantChecks.require(
          signatureLine.previewFormat() != null, "signature line previewFormat must exist");
    }
    if (signatureLine.previewByteSize() != null) {
      WorkbookInvariantChecks.require(
          signatureLine.previewByteSize() >= 0L,
          "signature line previewByteSize must not be negative");
      WorkbookInvariantChecks.require(
          signatureLine.previewFormat() != null, "signature line previewFormat must exist");
    }
    if (signatureLine.previewSha256() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.previewSha256(), "signature line previewSha256");
      WorkbookInvariantChecks.require(
          signatureLine.previewFormat() != null, "signature line previewFormat must exist");
    }
    if (signatureLine.previewWidthPixels() != null) {
      WorkbookInvariantChecks.require(
          signatureLine.previewWidthPixels() >= 0,
          "signature line previewWidthPixels must not be negative");
    }
    if (signatureLine.previewHeightPixels() != null) {
      WorkbookInvariantChecks.require(
          signatureLine.previewHeightPixels() >= 0,
          "signature line previewHeightPixels must not be negative");
    }
  }

  static void requireDrawingObjectPayloadShape(DrawingObjectPayloadReport payload) {
    WorkbookInvariantChecks.require(payload != null, "drawing payload must not be null");
    WorkbookInvariantChecks.requireNonBlank(payload.name(), "drawing payload name");
    WorkbookInvariantChecks.requireNonBlank(payload.contentType(), "drawing payload contentType");
    WorkbookInvariantChecks.requireNonBlank(payload.sha256(), "drawing payload sha256");
    WorkbookInvariantChecks.requireBase64(payload.base64Data(), "drawing payload base64Data");
    switch (payload) {
      case DrawingObjectPayloadReport.Picture picture -> {
        WorkbookInvariantChecks.requireNonBlank(picture.fileName(), "picture payload fileName");
        if (picture.description() != null) {
          WorkbookInvariantChecks.require(
              !picture.description().isBlank(), "picture payload description must not be blank");
        }
      }
      case DrawingObjectPayloadReport.EmbeddedObject embeddedObject -> {
        if (embeddedObject.fileName() != null) {
          WorkbookInvariantChecks.require(
              !embeddedObject.fileName().isBlank(), "embedded payload fileName must not be blank");
        }
        if (embeddedObject.label() != null) {
          WorkbookInvariantChecks.require(
              !embeddedObject.label().isBlank(), "embedded payload label must not be blank");
        }
        if (embeddedObject.command() != null) {
          WorkbookInvariantChecks.require(
              !embeddedObject.command().isBlank(), "embedded payload command must not be blank");
        }
      }
    }
  }

  static void requireDrawingAnchorShape(DrawingAnchorReport anchor) {
    WorkbookInvariantChecks.require(anchor != null, "drawing anchor must not be null");
    switch (anchor) {
      case DrawingAnchorReport.TwoCell twoCell -> {
        requireDrawingMarkerShape(twoCell.from());
        requireDrawingMarkerShape(twoCell.to());
      }
      case DrawingAnchorReport.OneCell oneCell -> {
        requireDrawingMarkerShape(oneCell.from());
        WorkbookInvariantChecks.require(
            oneCell.widthEmu() > 0L, "one-cell widthEmu must be positive");
        WorkbookInvariantChecks.require(
            oneCell.heightEmu() > 0L, "one-cell heightEmu must be positive");
      }
      case DrawingAnchorReport.Absolute absolute -> {
        WorkbookInvariantChecks.require(
            absolute.xEmu() >= 0L, "absolute xEmu must not be negative");
        WorkbookInvariantChecks.require(
            absolute.yEmu() >= 0L, "absolute yEmu must not be negative");
        WorkbookInvariantChecks.require(
            absolute.widthEmu() > 0L, "absolute widthEmu must be positive");
        WorkbookInvariantChecks.require(
            absolute.heightEmu() > 0L, "absolute heightEmu must be positive");
      }
    }
  }

  static void requireDrawingMarkerShape(DrawingMarkerReport marker) {
    WorkbookInvariantChecks.require(marker != null, "drawing marker must not be null");
    WorkbookInvariantChecks.require(
        marker.columnIndex() >= 0, "drawing marker columnIndex must not be negative");
    WorkbookInvariantChecks.require(
        marker.rowIndex() >= 0, "drawing marker rowIndex must not be negative");
    WorkbookInvariantChecks.require(marker.dx() >= 0, "drawing marker dx must not be negative");
    WorkbookInvariantChecks.require(marker.dy() >= 0, "drawing marker dy must not be negative");
  }

  static void requireChartReportShape(ChartReport chart) {
    WorkbookInvariantChecks.require(chart != null, "chart report must not be null");
    WorkbookInvariantChecks.requireNonBlank(chart.name(), "chart name");
    requireDrawingAnchorShape(chart.anchor());
    requireChartTitleShape(chart.title());
    requireChartLegendShape(chart.legend());
    WorkbookInvariantChecks.require(
        chart.displayBlanksAs() != null, "chart displayBlanksAs must not be null");
    chart.plots().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartPlotShape);
  }

  static void requireChartLegendShape(ChartReport.Legend legend) {
    WorkbookInvariantChecks.require(legend != null, "chart legend must not be null");
    switch (legend) {
      case ChartReport.Legend.Hidden _ -> {}
      case ChartReport.Legend.Visible visible ->
          WorkbookInvariantChecks.require(
              visible.position() != null, "chart legend position must not be null");
    }
  }

  static void requireChartPlotShape(ChartReport.Plot plot) {
    WorkbookInvariantChecks.require(plot != null, "chart plot must not be null");
    switch (plot) {
      case ChartReport.Area area -> {
        WorkbookInvariantChecks.require(
            area.grouping() != null, "area chart grouping must not be null");
        area.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        area.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Area3D area3D -> {
        WorkbookInvariantChecks.require(
            area3D.grouping() != null, "area 3D chart grouping must not be null");
        area3D.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        area3D.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Bar bar -> {
        WorkbookInvariantChecks.require(
            bar.barDirection() != null, "bar chart barDirection must not be null");
        WorkbookInvariantChecks.require(
            bar.grouping() != null, "bar chart grouping must not be null");
        bar.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        bar.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Bar3D bar3D -> {
        WorkbookInvariantChecks.require(
            bar3D.barDirection() != null, "bar 3D chart barDirection must not be null");
        WorkbookInvariantChecks.require(
            bar3D.grouping() != null, "bar 3D chart grouping must not be null");
        bar3D.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        bar3D.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Doughnut doughnut -> {
        if (doughnut.firstSliceAngle() != null) {
          WorkbookInvariantChecks.require(
              doughnut.firstSliceAngle() >= 0 && doughnut.firstSliceAngle() <= 360,
              "doughnut chart firstSliceAngle must be between 0 and 360");
        }
        if (doughnut.holeSize() != null) {
          WorkbookInvariantChecks.require(
              doughnut.holeSize() >= 10 && doughnut.holeSize() <= 90,
              "doughnut chart holeSize must be between 10 and 90");
        }
        doughnut.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Line line -> {
        WorkbookInvariantChecks.require(
            line.grouping() != null, "line chart grouping must not be null");
        line.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        line.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Line3D line3D -> {
        WorkbookInvariantChecks.require(
            line3D.grouping() != null, "line 3D chart grouping must not be null");
        line3D.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        line3D.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Pie pie -> {
        if (pie.firstSliceAngle() != null) {
          WorkbookInvariantChecks.require(
              pie.firstSliceAngle() >= 0 && pie.firstSliceAngle() <= 360,
              "pie chart firstSliceAngle must be between 0 and 360");
        }
        pie.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Pie3D pie3D ->
          pie3D.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      case ChartReport.Radar radar -> {
        WorkbookInvariantChecks.require(
            radar.style() != null, "radar chart style must not be null");
        radar.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        radar.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Scatter scatter -> {
        WorkbookInvariantChecks.require(
            scatter.style() != null, "scatter chart style must not be null");
        scatter.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        scatter.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Surface surface -> {
        surface.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        surface.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Surface3D surface3D -> {
        surface3D.axes().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartAxisShape);
        surface3D.series().forEach(WorkbookInvariantWorkbookSurfaceChecks::requireChartSeriesShape);
      }
      case ChartReport.Unsupported unsupported -> {
        WorkbookInvariantChecks.requireNonBlank(
            unsupported.plotTypeToken(), "unsupported chart plotTypeToken");
        WorkbookInvariantChecks.requireNonBlank(unsupported.detail(), "unsupported chart detail");
      }
    }
  }

  static void requireChartAxisShape(ChartReport.Axis axis) {
    WorkbookInvariantChecks.require(axis != null, "chart axis must not be null");
    WorkbookInvariantChecks.require(axis.kind() != null, "chart axis kind must not be null");
    WorkbookInvariantChecks.require(
        axis.position() != null, "chart axis position must not be null");
    WorkbookInvariantChecks.require(axis.crosses() != null, "chart axis crosses must not be null");
  }

  static void requireChartSeriesShape(ChartReport.Series series) {
    WorkbookInvariantChecks.require(series != null, "chart series must not be null");
    requireChartTitleShape(series.title());
    requireChartDataSourceShape(series.categories());
    requireChartDataSourceShape(series.values());
  }

  static void requireChartTitleShape(ChartReport.Title title) {
    WorkbookInvariantChecks.require(title != null, "chart title must not be null");
    switch (title) {
      case ChartReport.Title.None _ -> {}
      case ChartReport.Title.Text text ->
          WorkbookInvariantChecks.require(text.text() != null, "chart title text must not be null");
      case ChartReport.Title.Formula formula -> {
        WorkbookInvariantChecks.requireNonBlank(formula.formula(), "chart title formula");
        WorkbookInvariantChecks.require(
            formula.cachedText() != null, "chart title cachedText must not be null");
      }
    }
  }

  static void requireChartDataSourceShape(ChartReport.DataSource source) {
    WorkbookInvariantChecks.require(source != null, "chart data source must not be null");
    switch (source) {
      case ChartReport.DataSource.StringReference reference -> {
        WorkbookInvariantChecks.requireNonBlank(
            reference.formula(), "chart string-reference formula");
        WorkbookInvariantChecks.require(
            reference.cachedValues() != null,
            "chart string-reference cachedValues must not be null");
      }
      case ChartReport.DataSource.NumericReference reference -> {
        WorkbookInvariantChecks.requireNonBlank(
            reference.formula(), "chart numeric-reference formula");
        WorkbookInvariantChecks.require(
            reference.cachedValues() != null,
            "chart numeric-reference cachedValues must not be null");
        if (reference.formatCode() != null) {
          WorkbookInvariantChecks.require(
              !reference.formatCode().isBlank(),
              "chart numeric-reference formatCode must not be blank");
        }
      }
      case ChartReport.DataSource.StringLiteral literal ->
          WorkbookInvariantChecks.require(
              literal.values() != null, "chart string-literal values must not be null");
      case ChartReport.DataSource.NumericLiteral literal -> {
        WorkbookInvariantChecks.require(
            literal.values() != null, "chart numeric-literal values must not be null");
        if (literal.formatCode() != null) {
          WorkbookInvariantChecks.require(
              !literal.formatCode().isBlank(),
              "chart numeric-literal formatCode must not be blank");
        }
      }
    }
  }

  static void requirePivotTableShape(PivotTableReport pivotTable) {
    WorkbookInvariantChecks.require(pivotTable != null, "pivot table must not be null");
    WorkbookInvariantChecks.requireNonBlank(pivotTable.name(), "pivot table name");
    WorkbookInvariantChecks.requireNonBlank(pivotTable.sheetName(), "pivot table sheetName");
    requirePivotTableAnchorShape(pivotTable.anchor());
    switch (pivotTable) {
      case PivotTableReport.Supported supported -> {
        requirePivotTableSourceShape(supported.source());
        supported
            .rowLabels()
            .forEach(WorkbookInvariantWorkbookSurfaceChecks::requirePivotTableFieldShape);
        supported
            .columnLabels()
            .forEach(WorkbookInvariantWorkbookSurfaceChecks::requirePivotTableFieldShape);
        supported
            .reportFilters()
            .forEach(WorkbookInvariantWorkbookSurfaceChecks::requirePivotTableFieldShape);
        WorkbookInvariantChecks.require(
            !supported.dataFields().isEmpty(), "pivot table dataFields must not be empty");
        supported
            .dataFields()
            .forEach(WorkbookInvariantWorkbookSurfaceChecks::requirePivotTableDataFieldShape);
      }
      case PivotTableReport.Unsupported unsupported ->
          WorkbookInvariantChecks.requireNonBlank(unsupported.detail(), "pivot table detail");
    }
  }

  static void requirePivotTableSourceShape(PivotTableReport.Source source) {
    WorkbookInvariantChecks.require(source != null, "pivot table source must not be null");
    switch (source) {
      case PivotTableReport.Source.Range range -> {
        WorkbookInvariantChecks.requireNonBlank(range.sheetName(), "pivot range source sheetName");
        WorkbookInvariantChecks.requireNonBlank(range.range(), "pivot range source range");
      }
      case PivotTableReport.Source.NamedRange namedRange -> {
        WorkbookInvariantChecks.requireNonBlank(namedRange.name(), "pivot named-range source name");
        WorkbookInvariantChecks.requireNonBlank(
            namedRange.sheetName(), "pivot named-range source sheetName");
        WorkbookInvariantChecks.requireNonBlank(
            namedRange.range(), "pivot named-range source range");
      }
      case PivotTableReport.Source.Table table -> {
        WorkbookInvariantChecks.requireNonBlank(table.name(), "pivot table source name");
        WorkbookInvariantChecks.requireNonBlank(table.sheetName(), "pivot table source sheetName");
        WorkbookInvariantChecks.requireNonBlank(table.range(), "pivot table source range");
      }
    }
  }

  static void requirePivotTableAnchorShape(PivotTableReport.Anchor anchor) {
    WorkbookInvariantChecks.require(anchor != null, "pivot table anchor must not be null");
    WorkbookInvariantChecks.requireNonBlank(
        anchor.topLeftAddress(), "pivot table anchor topLeftAddress");
    WorkbookInvariantChecks.requireNonBlank(
        anchor.locationRange(), "pivot table anchor locationRange");
  }

  static void requirePivotTableFieldShape(PivotTableReport.Field field) {
    WorkbookInvariantChecks.require(field != null, "pivot table field must not be null");
    WorkbookInvariantChecks.require(
        field.sourceColumnIndex() >= 0, "pivot field sourceColumnIndex must not be negative");
    WorkbookInvariantChecks.requireNonBlank(
        field.sourceColumnName(), "pivot field sourceColumnName");
  }

  static void requirePivotTableDataFieldShape(PivotTableReport.DataField dataField) {
    WorkbookInvariantChecks.require(dataField != null, "pivot table dataField must not be null");
    WorkbookInvariantChecks.require(
        dataField.sourceColumnIndex() >= 0,
        "pivot dataField sourceColumnIndex must not be negative");
    WorkbookInvariantChecks.requireNonBlank(
        dataField.sourceColumnName(), "pivot dataField sourceColumnName");
    WorkbookInvariantChecks.require(
        dataField.function() != null, "pivot dataField function must not be null");
    WorkbookInvariantChecks.requireNonBlank(dataField.displayName(), "pivot dataField displayName");
    if (dataField.valueFormat() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          dataField.valueFormat(), "pivot dataField valueFormat");
    }
  }
}
