package dev.erst.gridgrind.jazzer.support;

import dev.erst.gridgrind.excel.*;
import java.util.HashSet;

/** Owns invariant checks for engine-native workbook, drawing, chart, and pivot-table snapshots. */
final class WorkbookInvariantEngineShapeChecks {
  private WorkbookInvariantEngineShapeChecks() {}

  static void requireEngineWorkbookSummaryShape(
      dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary workbook) {
    WorkbookInvariantChecks.require(workbook != null, "engine workbook summary must not be null");
    WorkbookInvariantChecks.require(
        workbook.sheetCount() >= 0, "engine sheetCount must not be negative");
    WorkbookInvariantChecks.require(
        workbook.namedRangeCount() >= 0, "engine namedRangeCount must not be negative");
    WorkbookInvariantChecks.require(
        workbook.sheetNames() != null, "engine sheetNames must not be null");
    WorkbookInvariantChecks.require(
        workbook.sheetCount() == workbook.sheetNames().size(),
        "engine sheetCount must match sheetNames size");
    WorkbookInvariantChecks.require(
        workbook.sheetNames().size() == new HashSet<>(workbook.sheetNames()).size(),
        "engine sheet names must be unique");
    workbook
        .sheetNames()
        .forEach(
            sheetName -> WorkbookInvariantChecks.requireNonBlank(sheetName, "engine sheetName"));
    switch (workbook) {
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.Empty empty -> {
        WorkbookInvariantChecks.require(
            empty.sheetCount() == 0, "engine empty workbook summary must have sheetCount 0");
        WorkbookInvariantChecks.require(
            empty.sheetNames().isEmpty(), "engine empty workbook summary must have no sheet names");
      }
      case dev.erst.gridgrind.excel.WorkbookCoreResult.WorkbookSummary.WithSheets withSheets -> {
        WorkbookInvariantChecks.require(
            withSheets.sheetCount() > 0,
            "engine non-empty workbook summary must have positive sheetCount");
        WorkbookInvariantChecks.requireNonBlank(
            withSheets.activeSheetName(), "engine activeSheetName");
        WorkbookInvariantChecks.require(
            withSheets.sheetNames().contains(withSheets.activeSheetName()),
            "engine activeSheetName must be present in sheetNames");
        WorkbookInvariantChecks.require(
            !withSheets.selectedSheetNames().isEmpty(),
            "engine selectedSheetNames must not be empty");
        WorkbookInvariantChecks.require(
            withSheets.selectedSheetNames().size()
                == new HashSet<>(withSheets.selectedSheetNames()).size(),
            "engine selectedSheetNames must be unique");
        withSheets
            .selectedSheetNames()
            .forEach(
                selectedSheetName -> {
                  WorkbookInvariantChecks.requireNonBlank(
                      selectedSheetName, "engine selectedSheetName");
                  WorkbookInvariantChecks.require(
                      withSheets.sheetNames().contains(selectedSheetName),
                      "engine selectedSheetNames must be present in sheetNames");
                });
      }
    }
  }

  static void requireEngineSheetSummaryShape(
      dev.erst.gridgrind.excel.WorkbookSheetResult.SheetSummary sheet) {
    WorkbookInvariantChecks.requireNonBlank(sheet.sheetName(), "engine sheetName");
    WorkbookInvariantChecks.require(
        sheet.visibility() != null, "engine visibility must not be null");
    WorkbookInvariantChecks.require(
        sheet.protection() != null, "engine protection must not be null");
    WorkbookInvariantChecks.require(
        sheet.physicalRowCount() >= 0, "engine physicalRowCount must not be negative");
    WorkbookInvariantChecks.require(
        sheet.lastRowIndex() >= -1, "engine lastRowIndex must be greater than or equal to -1");
    WorkbookInvariantChecks.require(
        sheet.lastColumnIndex() >= -1,
        "engine lastColumnIndex must be greater than or equal to -1");
    switch (sheet.protection()) {
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Unprotected _ -> {}
      case dev.erst.gridgrind.excel.WorkbookSheetResult.SheetProtection.Protected protectedSheet ->
          WorkbookInvariantChecks.require(
              protectedSheet.settings() != null,
              "engine protected sheet settings must not be null");
    }
  }

  static void requireEngineDrawingObjectShape(
      dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot drawingObject) {
    WorkbookInvariantChecks.require(
        drawingObject != null, "engine drawing object must not be null");
    WorkbookInvariantChecks.requireNonBlank(drawingObject.name(), "engine drawing object name");
    requireEngineDrawingAnchorShape(drawingObject.anchor());
    switch (drawingObject) {
      case dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.Picture picture -> {
        WorkbookInvariantChecks.requireNonBlank(
            picture.contentType(), "engine picture contentType");
        WorkbookInvariantChecks.requireNonBlank(picture.sha256(), "engine picture sha256");
        WorkbookInvariantChecks.require(
            picture.byteSize() >= 0L, "engine picture byteSize must not be negative");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.Chart chart -> {
        WorkbookInvariantChecks.require(
            chart.plotTypeTokens() != null, "engine chart plotTypeTokens must not be null");
        chart
            .plotTypeTokens()
            .forEach(
                token ->
                    WorkbookInvariantChecks.requireNonBlank(token, "engine chart plotTypeToken"));
        WorkbookInvariantChecks.require(
            chart.title() != null, "engine chart title must not be null");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.Shape shape -> {
        if (shape.presetGeometryToken() != null) {
          WorkbookInvariantChecks.require(
              !shape.presetGeometryToken().isBlank(),
              "engine shape presetGeometryToken must not be blank");
        }
        if (shape.text() != null) {
          WorkbookInvariantChecks.require(
              !shape.text().isBlank(), "engine shape text must not be blank");
        }
        WorkbookInvariantChecks.require(
            shape.childCount() >= 0, "engine shape childCount must not be negative");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.EmbeddedObject embeddedObject -> {
        WorkbookInvariantChecks.requireNonBlank(
            embeddedObject.contentType(), "engine embedded contentType");
        WorkbookInvariantChecks.requireNonBlank(embeddedObject.sha256(), "engine embedded sha256");
        WorkbookInvariantChecks.require(
            embeddedObject.byteSize() >= 0L, "engine embedded byteSize must not be negative");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.SignatureLine signatureLine ->
          requireEngineSignatureLineDrawingObjectShape(signatureLine);
    }
  }

  static void requireEngineChartShape(dev.erst.gridgrind.excel.ExcelChartSnapshot chart) {
    WorkbookInvariantChecks.require(chart != null, "engine chart must not be null");
    WorkbookInvariantChecks.requireNonBlank(chart.name(), "engine chart name");
    requireEngineDrawingAnchorShape(chart.anchor());
    requireEngineChartTitleShape(chart.title());
    requireEngineChartLegendShape(chart.legend());
    WorkbookInvariantChecks.require(
        chart.displayBlanksAs() != null, "engine chart displayBlanksAs must not be null");
    chart.plots().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartPlotShape);
  }

  static void requireEngineSignatureLineDrawingObjectShape(
      dev.erst.gridgrind.excel.ExcelDrawingObjectSnapshot.SignatureLine signatureLine) {
    if (signatureLine.setupId() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.setupId(), "engine signature line setupId");
    }
    if (signatureLine.signingInstructions() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.signingInstructions(), "engine signature line signingInstructions");
    }
    if (signatureLine.suggestedSigner() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.suggestedSigner(), "engine signature line suggestedSigner");
    }
    if (signatureLine.suggestedSigner2() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.suggestedSigner2(), "engine signature line suggestedSigner2");
    }
    if (signatureLine.suggestedSignerEmail() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.suggestedSignerEmail(), "engine signature line suggestedSignerEmail");
    }
    if (signatureLine.previewContentType() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.previewContentType(), "engine signature line previewContentType");
      WorkbookInvariantChecks.require(
          signatureLine.previewFormat() != null, "engine signature line previewFormat must exist");
    }
    if (signatureLine.previewByteSize() != null) {
      WorkbookInvariantChecks.require(
          signatureLine.previewByteSize() >= 0L,
          "engine signature line previewByteSize must not be negative");
      WorkbookInvariantChecks.require(
          signatureLine.previewFormat() != null, "engine signature line previewFormat must exist");
    }
    if (signatureLine.previewSha256() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          signatureLine.previewSha256(), "engine signature line previewSha256");
      WorkbookInvariantChecks.require(
          signatureLine.previewFormat() != null, "engine signature line previewFormat must exist");
    }
    if (signatureLine.previewWidthPixels() != null) {
      WorkbookInvariantChecks.require(
          signatureLine.previewWidthPixels() >= 0,
          "engine signature line previewWidthPixels must not be negative");
    }
    if (signatureLine.previewHeightPixels() != null) {
      WorkbookInvariantChecks.require(
          signatureLine.previewHeightPixels() >= 0,
          "engine signature line previewHeightPixels must not be negative");
    }
  }

  static void requireEngineChartLegendShape(
      dev.erst.gridgrind.excel.ExcelChartSnapshot.Legend legend) {
    WorkbookInvariantChecks.require(legend != null, "engine chart legend must not be null");
    switch (legend) {
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Legend.Hidden _ -> {}
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Legend.Visible visible ->
          WorkbookInvariantChecks.require(
              visible.position() != null, "engine chart legend position must not be null");
    }
  }

  static void requireEngineChartPlotShape(dev.erst.gridgrind.excel.ExcelChartSnapshot.Plot plot) {
    WorkbookInvariantChecks.require(plot != null, "engine chart plot must not be null");
    switch (plot) {
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Area area -> {
        WorkbookInvariantChecks.require(
            area.grouping() != null, "engine area chart grouping must not be null");
        area.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        area.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Area3D area3D -> {
        WorkbookInvariantChecks.require(
            area3D.grouping() != null, "engine area 3D chart grouping must not be null");
        area3D.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        area3D.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Bar bar -> {
        WorkbookInvariantChecks.require(
            bar.barDirection() != null, "engine bar chart barDirection must not be null");
        WorkbookInvariantChecks.require(
            bar.grouping() != null, "engine bar chart grouping must not be null");
        bar.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        bar.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Bar3D bar3D -> {
        WorkbookInvariantChecks.require(
            bar3D.barDirection() != null, "engine bar 3D chart barDirection must not be null");
        WorkbookInvariantChecks.require(
            bar3D.grouping() != null, "engine bar 3D chart grouping must not be null");
        bar3D.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        bar3D.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Doughnut doughnut -> {
        if (doughnut.firstSliceAngle() != null) {
          WorkbookInvariantChecks.require(
              doughnut.firstSliceAngle() >= 0 && doughnut.firstSliceAngle() <= 360,
              "engine doughnut chart firstSliceAngle must be between 0 and 360");
        }
        if (doughnut.holeSize() != null) {
          WorkbookInvariantChecks.require(
              doughnut.holeSize() >= 10 && doughnut.holeSize() <= 90,
              "engine doughnut chart holeSize must be between 10 and 90");
        }
        doughnut
            .series()
            .forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Line line -> {
        WorkbookInvariantChecks.require(
            line.grouping() != null, "engine line chart grouping must not be null");
        line.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        line.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Line3D line3D -> {
        WorkbookInvariantChecks.require(
            line3D.grouping() != null, "engine line 3D chart grouping must not be null");
        line3D.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        line3D.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Pie pie -> {
        if (pie.firstSliceAngle() != null) {
          WorkbookInvariantChecks.require(
              pie.firstSliceAngle() >= 0 && pie.firstSliceAngle() <= 360,
              "engine pie chart firstSliceAngle must be between 0 and 360");
        }
        pie.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Pie3D pie3D ->
          pie3D.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Radar radar -> {
        WorkbookInvariantChecks.require(
            radar.style() != null, "engine radar chart style must not be null");
        radar.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        radar.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Scatter scatter -> {
        WorkbookInvariantChecks.require(
            scatter.style() != null, "engine scatter chart style must not be null");
        scatter.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        scatter.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Surface surface -> {
        surface.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        surface.series().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Surface3D surface3D -> {
        surface3D.axes().forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartAxisShape);
        surface3D
            .series()
            .forEach(WorkbookInvariantEngineShapeChecks::requireEngineChartSeriesShape);
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Unsupported unsupported -> {
        WorkbookInvariantChecks.requireNonBlank(
            unsupported.plotTypeToken(), "engine unsupported chart plotTypeToken");
        WorkbookInvariantChecks.requireNonBlank(
            unsupported.detail(), "engine unsupported chart detail");
      }
    }
  }

  static void requireEnginePivotTableShape(
      dev.erst.gridgrind.excel.ExcelPivotTableSnapshot pivotTable) {
    WorkbookInvariantChecks.require(pivotTable != null, "engine pivot table must not be null");
    WorkbookInvariantChecks.requireNonBlank(pivotTable.name(), "engine pivot table name");
    WorkbookInvariantChecks.requireNonBlank(pivotTable.sheetName(), "engine pivot table sheetName");
    requireEnginePivotTableAnchorShape(pivotTable.anchor());
    switch (pivotTable) {
      case dev.erst.gridgrind.excel.ExcelPivotTableSnapshot.Supported supported -> {
        requireEnginePivotTableSourceShape(supported.source());
        supported
            .rowLabels()
            .forEach(WorkbookInvariantEngineShapeChecks::requireEnginePivotTableFieldShape);
        supported
            .columnLabels()
            .forEach(WorkbookInvariantEngineShapeChecks::requireEnginePivotTableFieldShape);
        supported
            .reportFilters()
            .forEach(WorkbookInvariantEngineShapeChecks::requireEnginePivotTableFieldShape);
        WorkbookInvariantChecks.require(
            !supported.dataFields().isEmpty(), "engine pivot table dataFields must not be empty");
        supported
            .dataFields()
            .forEach(WorkbookInvariantEngineShapeChecks::requireEnginePivotTableDataFieldShape);
      }
      case dev.erst.gridgrind.excel.ExcelPivotTableSnapshot.Unsupported unsupported ->
          WorkbookInvariantChecks.requireNonBlank(
              unsupported.detail(), "engine pivot table detail");
    }
  }

  static void requireEnginePivotTableSourceShape(
      dev.erst.gridgrind.excel.ExcelPivotTableSnapshot.Source source) {
    WorkbookInvariantChecks.require(source != null, "engine pivot table source must not be null");
    switch (source) {
      case dev.erst.gridgrind.excel.ExcelPivotTableSnapshot.Source.Range range -> {
        WorkbookInvariantChecks.requireNonBlank(
            range.sheetName(), "engine pivot range source sheetName");
        WorkbookInvariantChecks.requireNonBlank(range.range(), "engine pivot range source range");
      }
      case dev.erst.gridgrind.excel.ExcelPivotTableSnapshot.Source.NamedRange namedRange -> {
        WorkbookInvariantChecks.requireNonBlank(
            namedRange.name(), "engine pivot named-range source name");
        WorkbookInvariantChecks.requireNonBlank(
            namedRange.sheetName(), "engine pivot named-range source sheetName");
        WorkbookInvariantChecks.requireNonBlank(
            namedRange.range(), "engine pivot named-range source range");
      }
      case dev.erst.gridgrind.excel.ExcelPivotTableSnapshot.Source.Table table -> {
        WorkbookInvariantChecks.requireNonBlank(table.name(), "engine pivot table source name");
        WorkbookInvariantChecks.requireNonBlank(
            table.sheetName(), "engine pivot table source sheetName");
        WorkbookInvariantChecks.requireNonBlank(table.range(), "engine pivot table source range");
      }
    }
  }

  static void requireEnginePivotTableAnchorShape(
      dev.erst.gridgrind.excel.ExcelPivotTableSnapshot.Anchor anchor) {
    WorkbookInvariantChecks.require(anchor != null, "engine pivot table anchor must not be null");
    WorkbookInvariantChecks.requireNonBlank(
        anchor.topLeftAddress(), "engine pivot table anchor topLeftAddress");
    WorkbookInvariantChecks.requireNonBlank(
        anchor.locationRange(), "engine pivot table anchor locationRange");
  }

  static void requireEnginePivotTableFieldShape(
      dev.erst.gridgrind.excel.ExcelPivotTableSnapshot.Field field) {
    WorkbookInvariantChecks.require(field != null, "engine pivot table field must not be null");
    WorkbookInvariantChecks.require(
        field.sourceColumnIndex() >= 0,
        "engine pivot field sourceColumnIndex must not be negative");
    WorkbookInvariantChecks.requireNonBlank(
        field.sourceColumnName(), "engine pivot field sourceColumnName");
  }

  static void requireEnginePivotTableDataFieldShape(
      dev.erst.gridgrind.excel.ExcelPivotTableSnapshot.DataField dataField) {
    WorkbookInvariantChecks.require(
        dataField != null, "engine pivot table dataField must not be null");
    WorkbookInvariantChecks.require(
        dataField.sourceColumnIndex() >= 0,
        "engine pivot dataField sourceColumnIndex must not be negative");
    WorkbookInvariantChecks.requireNonBlank(
        dataField.sourceColumnName(), "engine pivot dataField sourceColumnName");
    WorkbookInvariantChecks.require(
        dataField.function() != null, "engine pivot dataField function must not be null");
    WorkbookInvariantChecks.requireNonBlank(
        dataField.displayName(), "engine pivot dataField displayName");
    if (dataField.valueFormat() != null) {
      WorkbookInvariantChecks.requireNonBlank(
          dataField.valueFormat(), "engine pivot dataField valueFormat");
    }
  }

  static void requireEngineChartAxisShape(dev.erst.gridgrind.excel.ExcelChartSnapshot.Axis axis) {
    WorkbookInvariantChecks.require(axis != null, "engine chart axis must not be null");
    WorkbookInvariantChecks.require(axis.kind() != null, "engine chart axis kind must not be null");
    WorkbookInvariantChecks.require(
        axis.position() != null, "engine chart axis position must not be null");
    WorkbookInvariantChecks.require(
        axis.crosses() != null, "engine chart axis crosses must not be null");
  }

  static void requireEngineChartSeriesShape(
      dev.erst.gridgrind.excel.ExcelChartSnapshot.Series series) {
    WorkbookInvariantChecks.require(series != null, "engine chart series must not be null");
    requireEngineChartTitleShape(series.title());
    requireEngineChartDataSourceShape(series.categories());
    requireEngineChartDataSourceShape(series.values());
  }

  static void requireEngineChartTitleShape(
      dev.erst.gridgrind.excel.ExcelChartSnapshot.Title title) {
    WorkbookInvariantChecks.require(title != null, "engine chart title must not be null");
    switch (title) {
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Title.None _ -> {}
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Title.Text text ->
          WorkbookInvariantChecks.require(
              text.text() != null, "engine chart title text must not be null");
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.Title.Formula formula -> {
        WorkbookInvariantChecks.requireNonBlank(formula.formula(), "engine chart title formula");
        WorkbookInvariantChecks.require(
            formula.cachedText() != null, "engine chart title cachedText must not be null");
      }
    }
  }

  static void requireEngineChartDataSourceShape(
      dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource source) {
    WorkbookInvariantChecks.require(source != null, "engine chart data source must not be null");
    switch (source) {
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource.StringReference reference -> {
        WorkbookInvariantChecks.requireNonBlank(
            reference.formula(), "engine chart string-reference formula");
        WorkbookInvariantChecks.require(
            reference.cachedValues() != null,
            "engine chart string-reference cachedValues must not be null");
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource.NumericReference reference -> {
        WorkbookInvariantChecks.requireNonBlank(
            reference.formula(), "engine chart numeric-reference formula");
        WorkbookInvariantChecks.require(
            reference.cachedValues() != null,
            "engine chart numeric-reference cachedValues must not be null");
        if (reference.formatCode() != null) {
          WorkbookInvariantChecks.require(
              !reference.formatCode().isBlank(),
              "engine chart numeric-reference formatCode must not be blank");
        }
      }
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource.StringLiteral literal ->
          WorkbookInvariantChecks.require(
              literal.values() != null, "engine chart string-literal values must not be null");
      case dev.erst.gridgrind.excel.ExcelChartSnapshot.DataSource.NumericLiteral literal -> {
        WorkbookInvariantChecks.require(
            literal.values() != null, "engine chart numeric-literal values must not be null");
        if (literal.formatCode() != null) {
          WorkbookInvariantChecks.require(
              !literal.formatCode().isBlank(),
              "engine chart numeric-literal formatCode must not be blank");
        }
      }
    }
  }

  static void requireEngineDrawingAnchorShape(dev.erst.gridgrind.excel.ExcelDrawingAnchor anchor) {
    WorkbookInvariantChecks.require(anchor != null, "engine drawing anchor must not be null");
    switch (anchor) {
      case dev.erst.gridgrind.excel.ExcelDrawingAnchor.TwoCell twoCell -> {
        requireEngineDrawingMarkerShape(twoCell.from());
        requireEngineDrawingMarkerShape(twoCell.to());
        WorkbookInvariantChecks.require(
            twoCell.behavior() != null, "engine two-cell anchor behavior must not be null");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingAnchor.OneCell oneCell -> {
        requireEngineDrawingMarkerShape(oneCell.from());
        WorkbookInvariantChecks.require(
            oneCell.widthEmu() > 0L, "engine one-cell widthEmu must be positive");
        WorkbookInvariantChecks.require(
            oneCell.heightEmu() > 0L, "engine one-cell heightEmu must be positive");
        WorkbookInvariantChecks.require(
            oneCell.behavior() != null, "engine one-cell anchor behavior must not be null");
      }
      case dev.erst.gridgrind.excel.ExcelDrawingAnchor.Absolute absolute -> {
        WorkbookInvariantChecks.require(
            absolute.xEmu() >= 0L, "engine absolute xEmu must not be negative");
        WorkbookInvariantChecks.require(
            absolute.yEmu() >= 0L, "engine absolute yEmu must not be negative");
        WorkbookInvariantChecks.require(
            absolute.widthEmu() > 0L, "engine absolute widthEmu must be positive");
        WorkbookInvariantChecks.require(
            absolute.heightEmu() > 0L, "engine absolute heightEmu must be positive");
        WorkbookInvariantChecks.require(
            absolute.behavior() != null, "engine absolute anchor behavior must not be null");
      }
    }
  }

  static void requireEngineDrawingMarkerShape(dev.erst.gridgrind.excel.ExcelDrawingMarker marker) {
    WorkbookInvariantChecks.require(marker != null, "engine drawing marker must not be null");
    WorkbookInvariantChecks.require(
        marker.columnIndex() >= 0, "engine drawing marker columnIndex must not be negative");
    WorkbookInvariantChecks.require(
        marker.rowIndex() >= 0, "engine drawing marker rowIndex must not be negative");
    WorkbookInvariantChecks.require(
        marker.dx() >= 0, "engine drawing marker dx must not be negative");
    WorkbookInvariantChecks.require(
        marker.dy() >= 0, "engine drawing marker dy must not be negative");
  }
}
