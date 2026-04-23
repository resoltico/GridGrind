package dev.erst.gridgrind.executor.parity;

import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.inspect;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.mutate;
import static dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.*;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.ProbeContext;
import dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.ProbeResult;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Drawing, chart, and pivot parity probes. */
final class XlsxParityDrawingChartPivotProbeGroup {
  private XlsxParityDrawingChartPivotProbeGroup() {}

  static ProbeResult probeDrawingReadback(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario drawing =
        context.scenario(XlsxParityScenarios.DRAWING_IMAGE);
    XlsxParityOracle.DrawingSheetSnapshot direct =
        XlsxParityOracle.drawingSheet(drawing.workbookPath(), "Ops");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            drawing.workbookPath(),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Ops"),
                new InspectionQuery.GetDrawingObjects()),
            inspect(
                "payload",
                new DrawingObjectSelector.ByName("Ops", "OpsPicture"),
                new InspectionQuery.GetDrawingObjectPayload()));
    InspectionResult.DrawingObjectsResult drawingObjects =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    InspectionResult.DrawingObjectPayloadResult payload =
        XlsxParityGridGrind.read(
            success, "payload", InspectionResult.DrawingObjectPayloadResult.class);
    if (!XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECT_PAYLOAD")) {
      return fail("GridGrind is missing the Phase 5 drawing read types.");
    }
    if (direct.objects().size() != 1
        || !(direct.objects().getFirst()
            instanceof XlsxParityOracle.PictureDrawingObjectSnapshot picture)
        || !direct.mergedRegions().isEmpty()
        || !direct.comments().isEmpty()) {
      return fail("Phase 5 drawing corpus workbook is not in the expected picture-only shape.");
    }
    if (drawingObjects.drawingObjects().size() != 1
        || !(drawingObjects.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Picture reportPicture)
        || !(payload.payload() instanceof DrawingObjectPayloadReport.Picture picturePayload)) {
      return fail("GridGrind did not return the expected picture drawing reports.");
    }
    boolean matches =
        "OpsPicture".equals(reportPicture.name())
            && matchesAnchor(reportPicture.anchor(), picture.anchor())
            && picture.pictureDigest().equals(reportPicture.sha256())
            && picture.pictureDigest().equals(picturePayload.sha256());
    return matches
        ? pass("GridGrind reads existing picture-backed drawing objects and payloads losslessly.")
        : fail("GridGrind drawing readback diverged from the direct POI picture oracle.");
  }

  static ProbeResult probeDrawingAuthoring(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario drawing =
        context.scenario(XlsxParityScenarios.DRAWING_AUTHORING);
    XlsxParityOracle.DrawingSheetSnapshot direct =
        XlsxParityOracle.drawingSheet(drawing.workbookPath(), "Ops");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            drawing.workbookPath(),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Ops"),
                new InspectionQuery.GetDrawingObjects()),
            inspect(
                "picture-payload",
                new DrawingObjectSelector.ByName("Ops", "OpsPicture"),
                new InspectionQuery.GetDrawingObjectPayload()),
            inspect(
                "embed-payload",
                new DrawingObjectSelector.ByName("Ops", "OpsEmbed"),
                new InspectionQuery.GetDrawingObjectPayload()));
    InspectionResult.DrawingObjectsResult drawingObjects =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    InspectionResult.DrawingObjectPayloadResult picturePayload =
        XlsxParityGridGrind.read(
            success, "picture-payload", InspectionResult.DrawingObjectPayloadResult.class);
    InspectionResult.DrawingObjectPayloadResult embedPayload =
        XlsxParityGridGrind.read(
            success, "embed-payload", InspectionResult.DrawingObjectPayloadResult.class);
    if (!XlsxParityGridGrind.hasMutationActionType("SET_PICTURE")
        || !XlsxParityGridGrind.hasMutationActionType("SET_SHAPE")
        || !XlsxParityGridGrind.hasMutationActionType("SET_EMBEDDED_OBJECT")
        || !XlsxParityGridGrind.hasMutationActionType("SET_DRAWING_OBJECT_ANCHOR")
        || !XlsxParityGridGrind.hasMutationActionType("DELETE_DRAWING_OBJECT")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECT_PAYLOAD")) {
      return fail("GridGrind is missing the Phase 5 drawing mutation or read contract.");
    }
    if (!direct.mergedRegions().isEmpty() || !direct.comments().isEmpty()) {
      return fail("GridGrind-authored Phase 5 drawing workbook contains unexpected extra state.");
    }
    List<String> directNames =
        direct.objects().stream().map(XlsxParityOracle.DirectDrawingObjectSnapshot::name).toList();
    List<String> reportedNames =
        drawingObjects.drawingObjects().stream().map(DrawingObjectReport::name).toList();
    if (!directNames.equals(List.of("OpsPicture", "OpsShape", "OpsEmbed"))
        || !directNames.equals(reportedNames)) {
      return fail(
          "GridGrind-authored drawing workbook did not persist the expected final drawing order."
              + " direct="
              + directNames
              + " reported="
              + reportedNames);
    }

    XlsxParityOracle.PictureDrawingObjectSnapshot directPicture =
        directObject(direct, "OpsPicture", XlsxParityOracle.PictureDrawingObjectSnapshot.class);
    XlsxParityOracle.ShapeDrawingObjectSnapshot directShape =
        directObject(direct, "OpsShape", XlsxParityOracle.ShapeDrawingObjectSnapshot.class);
    XlsxParityOracle.EmbeddedObjectDrawingObjectSnapshot directEmbedded =
        directObject(
            direct, "OpsEmbed", XlsxParityOracle.EmbeddedObjectDrawingObjectSnapshot.class);
    DrawingObjectReport.Picture reportPicture =
        drawingObjectReport(drawingObjects, "OpsPicture", DrawingObjectReport.Picture.class);
    DrawingObjectReport.Shape reportShape =
        drawingObjectReport(drawingObjects, "OpsShape", DrawingObjectReport.Shape.class);
    DrawingObjectReport.EmbeddedObject reportEmbedded =
        drawingObjectReport(drawingObjects, "OpsEmbed", DrawingObjectReport.EmbeddedObject.class);
    if (!(picturePayload.payload()
            instanceof DrawingObjectPayloadReport.Picture picturePayloadReport)
        || !(embedPayload.payload()
            instanceof DrawingObjectPayloadReport.EmbeddedObject embedPayloadReport)) {
      return fail("GridGrind-authored drawing payload reads returned unexpected report types.");
    }

    boolean matches =
        matchesAnchor(reportPicture.anchor(), directPicture.anchor())
            && directPicture.pictureDigest().equals(reportPicture.sha256())
            && directPicture.pictureDigest().equals(picturePayloadReport.sha256())
            && matchesAnchor(reportShape.anchor(), directShape.anchor())
            && "SIMPLE_SHAPE".equals(directShape.kind())
            && reportShape.kind()
                == dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind.SIMPLE_SHAPE
            && Objects.equals(directShape.presetGeometryToken(), reportShape.presetGeometryToken())
            && Objects.equals(directShape.text(), reportShape.text())
            && matchesAnchor(reportEmbedded.anchor(), directEmbedded.anchor())
            && directEmbedded.objectDigest().equals(reportEmbedded.sha256())
            && directEmbedded.objectDigest().equals(embedPayloadReport.sha256())
            && Objects.equals(directEmbedded.fileName(), reportEmbedded.fileName());
    return matches
        ? pass(
            "GridGrind authors, mutates, persists, and rereads Phase 5 drawing objects with direct POI agreement.")
        : fail("GridGrind-authored drawing workbook diverged from the direct POI oracle.");
  }

  static ProbeResult probeDrawingCommentCoexistence(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario drawing =
        context.copiedScenario(XlsxParityScenarios.DRAWING_COMMENTS, "drawing-comments-source");
    XlsxParityOracle.DrawingSheetSnapshot before =
        XlsxParityOracle.drawingSheet(drawing.workbookPath(), "Ops");
    Path outputPath = context.derivedWorkbook("drawing-comments-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        drawing.workbookPath(),
        outputPath,
        List.of(
            mutate(
                new CellSelector.ByAddress("Ops", "B2"),
                new MutationAction.SetComment(comment("Transient", "GridGrind", false))),
            mutate(new CellSelector.ByAddress("Ops", "B2"), new MutationAction.ClearComment())));
    XlsxParityOracle.DrawingSheetSnapshot after = XlsxParityOracle.drawingSheet(outputPath, "Ops");
    return before.equals(after)
        ? pass(
            "Comment operations coexist with drawing parts without corrupting pictures or comments.")
        : fail("Comment operations changed drawing-backed or comment-backed sheet state.");
  }

  static ProbeResult probeDrawingMergedImagePreservation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario drawing =
        context.copiedScenario(
            XlsxParityScenarios.DRAWING_MERGED_IMAGE, "drawing-merged-image-source");
    XlsxParityOracle.DrawingSheetSnapshot before =
        XlsxParityOracle.drawingSheet(drawing.workbookPath(), "Ops");
    Path outputPath = context.derivedWorkbook("drawing-merged-image-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        drawing.workbookPath(),
        outputPath,
        List.of(
            mutate(
                new CellSelector.ByAddress("Ops", "F8"),
                new MutationAction.SetCell(text("Touch")))));
    XlsxParityOracle.DrawingSheetSnapshot after = XlsxParityOracle.drawingSheet(outputPath, "Ops");
    return !before.mergedRegions().isEmpty() && before.equals(after)
        ? pass("Unrelated edits preserve pictures that coexist with merged regions.")
        : fail("GridGrind changed merged-region or picture-backed drawing state.");
  }

  static ProbeResult probeEmbeddedObjectPlatform(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario embedded =
        context.copiedScenario(XlsxParityScenarios.EMBEDDED_OBJECT, "embedded-object-source");
    XlsxParityOracle.DrawingSheetSnapshot before =
        XlsxParityOracle.drawingSheet(embedded.workbookPath(), "Objects");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            embedded.workbookPath(),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Objects"),
                new InspectionQuery.GetDrawingObjects()),
            inspect(
                "payload",
                new DrawingObjectSelector.ByName("Objects", "OpsEmbed"),
                new InspectionQuery.GetDrawingObjectPayload()));
    InspectionResult.DrawingObjectsResult drawingObjects =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    InspectionResult.DrawingObjectPayloadResult payload =
        XlsxParityGridGrind.read(
            success, "payload", InspectionResult.DrawingObjectPayloadResult.class);
    Path outputPath = context.derivedWorkbook("embedded-object-preserved");
    XlsxParityGridGrind.mutateWorkbook(
        embedded.workbookPath(),
        outputPath,
        List.of(
            mutate(
                new CellSelector.ByAddress("Objects", "G1"),
                new MutationAction.SetCell(text("Touch")))));
    XlsxParityOracle.DrawingSheetSnapshot after =
        XlsxParityOracle.drawingSheet(outputPath, "Objects");
    if (!XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECT_PAYLOAD")) {
      return fail("GridGrind is missing the Phase 5 embedded-object read types.");
    }
    if (before.objects().size() != 1
        || !(before.objects().getFirst()
            instanceof XlsxParityOracle.EmbeddedObjectDrawingObjectSnapshot directEmbedded)
        || !(drawingObjects.drawingObjects().getFirst()
            instanceof DrawingObjectReport.EmbeddedObject reportEmbedded)
        || !(payload.payload()
            instanceof DrawingObjectPayloadReport.EmbeddedObject payloadEmbedded)) {
      return fail("Embedded-object parity corpus did not surface the expected single OLE object.");
    }
    boolean readMatches =
        "OpsEmbed".equals(reportEmbedded.name())
            && matchesAnchor(reportEmbedded.anchor(), directEmbedded.anchor())
            && directEmbedded.objectDigest().equals(reportEmbedded.sha256())
            && directEmbedded.objectDigest().equals(payloadEmbedded.sha256())
            && Objects.equals(directEmbedded.fileName(), reportEmbedded.fileName());
    return readMatches && before.equals(after)
        ? pass("GridGrind reads and preserves embedded OLE payloads without package drift.")
        : fail("GridGrind embedded-object parity diverged from the direct POI oracle.");
  }

  static ProbeResult probeChartReadback(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario chart = context.scenario(XlsxParityScenarios.CHART);
    List<ChartReport> directCharts = XlsxParityOracle.charts(chart.workbookPath(), "Chart");
    XlsxParityOracle.DrawingSheetSnapshot directDrawing =
        XlsxParityOracle.drawingSheet(chart.workbookPath(), "Chart");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            chart.workbookPath(),
            inspect(
                "charts", new ChartSelector.AllOnSheet("Chart"), new InspectionQuery.GetCharts()),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Chart"),
                new InspectionQuery.GetDrawingObjects()));
    InspectionResult.ChartsResult charts =
        XlsxParityGridGrind.read(success, "charts", InspectionResult.ChartsResult.class);
    InspectionResult.DrawingObjectsResult drawing =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    if (!XlsxParityGridGrind.hasInspectionQueryType("GET_CHARTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")) {
      return fail("GridGrind is missing the Phase 6 chart read types.");
    }
    if (directCharts.size() != 1
        || !hasSinglePlot(directCharts.getFirst(), ChartReport.Bar.class)
        || directDrawing.objects().size() != 1
        || !(directDrawing.objects().getFirst()
            instanceof XlsxParityOracle.ChartDrawingObjectSnapshot directDrawingChart)) {
      return fail(
          "Phase 6 simple-chart corpus workbook is not in the expected single-chart shape.");
    }
    if (charts.charts().size() != 1
        || drawing.drawingObjects().size() != 1
        || !(drawing.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Chart reportChart)) {
      return fail("GridGrind did not return the expected chart readback shape.");
    }
    boolean matches =
        charts.charts().equals(directCharts)
            && chartDrawingMatches(reportChart, directDrawingChart)
            && drawing.drawingObjects().size() == 1;
    return matches
        ? pass(
            "GridGrind reads supported POI-authored simple charts and chart drawing inventory losslessly.")
        : fail("GridGrind chart readback diverged from the direct POI oracle.");
  }

  static ProbeResult probeChartAuthoring(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario chart =
        context.scenario(XlsxParityScenarios.CHART_AUTHORING);
    List<ChartReport> directCharts = XlsxParityOracle.charts(chart.workbookPath(), "Chart");
    XlsxParityOracle.DrawingSheetSnapshot directDrawing =
        XlsxParityOracle.drawingSheet(chart.workbookPath(), "Chart");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            chart.workbookPath(),
            inspect(
                "charts", new ChartSelector.AllOnSheet("Chart"), new InspectionQuery.GetCharts()),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Chart"),
                new InspectionQuery.GetDrawingObjects()));
    InspectionResult.ChartsResult charts =
        XlsxParityGridGrind.read(success, "charts", InspectionResult.ChartsResult.class);
    InspectionResult.DrawingObjectsResult drawing =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    if (!XlsxParityGridGrind.hasMutationActionType("SET_CHART")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_CHARTS")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_DRAWING_OBJECTS")) {
      return fail("GridGrind is missing the Phase 6 chart authoring contract.");
    }
    if (directCharts.size() != 1
        || !hasSinglePlot(directCharts.getFirst(), ChartReport.Bar.class)
        || directDrawing.objects().size() != 1
        || !(directDrawing.objects().getFirst()
            instanceof XlsxParityOracle.ChartDrawingObjectSnapshot directDrawingChart)) {
      return fail(
          "GridGrind-authored chart corpus workbook is not in the expected single-bar-chart shape.");
    }
    ChartReport directChart = directCharts.getFirst();
    ChartReport.Bar directChartPlot = onlyPlot(directChart, ChartReport.Bar.class);
    if (directChartPlot.series().size() != 2
        || !(directChartPlot.series().get(1).categories()
            instanceof ChartReport.DataSource.StringReference categories)
        || !(directChartPlot.series().get(1).values()
            instanceof ChartReport.DataSource.NumericReference values)
        || !"ChartCategories".equals(categories.formula())
        || !"ChartActual".equals(values.formula())) {
      return fail(
          "GridGrind-authored chart corpus workbook did not retain named-range-backed series binding.");
    }
    if (charts.charts().size() != 1
        || drawing.drawingObjects().size() != 1
        || !(drawing.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Chart reportChart)) {
      return fail("GridGrind-authored chart workbook did not surface the expected chart reports.");
    }
    boolean matches =
        charts.charts().equals(directCharts)
            && chartDrawingMatches(reportChart, directDrawingChart)
            && drawing.drawingObjects().size() == 1;
    return matches
        ? pass(
            "GridGrind authors named-range-backed simple charts and rereads them with direct POI agreement.")
        : fail("GridGrind-authored chart workbook diverged from the direct POI oracle.");
  }

  static ProbeResult probeChartMutation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario chart =
        context.copiedScenario(XlsxParityScenarios.CHART, "chart-mutation-source");
    Path outputPath = context.derivedWorkbook("chart-mutated");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.mutateWorkbook(
            chart.workbookPath(),
            outputPath,
            List.of(
                mutate(
                    new CellSelector.ByAddress("Chart", "F1"),
                    new MutationAction.SetCell(text("Touch"))),
                mutate(
                    new SheetSelector.ByName("Chart"),
                    new MutationAction.SetChart(
                        new ChartInput(
                            "OpsChart",
                            twoCellAnchorInput(
                                4, 1, 11, 17, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                            chartTitle("Actual focus"),
                            new ChartInput.Legend.Hidden(),
                            ExcelChartDisplayBlanksAs.ZERO,
                            true,
                            List.of(
                                new ChartInput.Bar(
                                    false,
                                    ExcelChartBarDirection.BAR,
                                    null,
                                    null,
                                    null,
                                    null,
                                    List.of(
                                        new ChartInput.Series(
                                            new ChartInput.Title.Formula("C1"),
                                            new ChartInput.DataSource.Reference("A2:A4"),
                                            new ChartInput.DataSource.Reference("C2:C4"),
                                            null,
                                            null,
                                            null,
                                            null)))))))),
            inspect(
                "charts", new ChartSelector.AllOnSheet("Chart"), new InspectionQuery.GetCharts()),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Chart"),
                new InspectionQuery.GetDrawingObjects()));
    List<ChartReport> directCharts = XlsxParityOracle.charts(outputPath, "Chart");
    XlsxParityOracle.DrawingSheetSnapshot directDrawing =
        XlsxParityOracle.drawingSheet(outputPath, "Chart");
    InspectionResult.ChartsResult charts =
        XlsxParityGridGrind.read(success, "charts", InspectionResult.ChartsResult.class);
    InspectionResult.DrawingObjectsResult drawing =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    if (directCharts.size() != 1
        || !hasSinglePlot(directCharts.getFirst(), ChartReport.Bar.class)
        || directDrawing.objects().size() != 1
        || !(directDrawing.objects().getFirst()
            instanceof XlsxParityOracle.ChartDrawingObjectSnapshot directDrawingChart)
        || charts.charts().size() != 1
        || drawing.drawingObjects().size() != 1
        || !(drawing.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Chart reportChart)) {
      return fail(
          "Chart mutation parity did not produce the expected single updated supported chart.");
    }
    ChartReport directChart = directCharts.getFirst();
    ChartReport.Bar directChartPlot = onlyPlot(directChart, ChartReport.Bar.class);
    boolean matches =
        charts.charts().equals(directCharts)
            && chartDrawingMatches(reportChart, directDrawingChart)
            && directChartPlot.barDirection() == ExcelChartBarDirection.BAR
            && directChart.plotOnlyVisibleCells()
            && !directChartPlot.varyColors()
            && directChart.legend() instanceof ChartReport.Legend.Hidden
            && directChart.title().equals(new ChartReport.Title.Text("Actual focus"));
    return matches
        ? pass(
            "GridGrind mutates existing simple charts after workbook-core edits with direct POI agreement.")
        : fail("GridGrind simple-chart mutation diverged from the direct POI oracle.");
  }

  static ProbeResult probeChartPreservation(ProbeContext context) {
    List<String> mismatches = new ArrayList<>();
    for (String scenarioId :
        List.of(
            XlsxParityScenarios.CHART,
            XlsxParityScenarios.CHART_AUTHORING,
            XlsxParityScenarios.CHART_UNSUPPORTED)) {
      XlsxParityScenarios.MaterializedScenario chart =
          context.copiedScenario(scenarioId, "chart-preservation-" + scenarioId);
      List<ChartReport> beforeCharts = XlsxParityOracle.charts(chart.workbookPath(), "Chart");
      XlsxParityOracle.DrawingSheetSnapshot beforeDrawing =
          XlsxParityOracle.drawingSheet(chart.workbookPath(), "Chart");
      Path outputPath = context.derivedWorkbook("chart-preserved-" + scenarioId.replace('-', '_'));
      XlsxParityGridGrind.mutateWorkbook(
          chart.workbookPath(),
          outputPath,
          List.of(
              mutate(
                  new CellSelector.ByAddress("Chart", "F8"),
                  new MutationAction.SetCell(text("Touch")))));
      List<ChartReport> afterCharts = XlsxParityOracle.charts(outputPath, "Chart");
      XlsxParityOracle.DrawingSheetSnapshot afterDrawing =
          XlsxParityOracle.drawingSheet(outputPath, "Chart");
      if (!beforeCharts.equals(afterCharts) || !beforeDrawing.equals(afterDrawing)) {
        mismatches.add(scenarioId);
      }
    }
    return mismatches.isEmpty()
        ? pass(
            "All Phase 6 chart corpus workbooks preserve chart relations and anchors across unrelated edits.")
        : fail("Unrelated GridGrind edits changed chart state for scenarios: " + mismatches);
  }

  static ProbeResult probeChartUnsupported(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario chart =
        context.copiedScenario(XlsxParityScenarios.CHART_UNSUPPORTED, "chart-unsupported-source");
    List<ChartReport> directCharts = XlsxParityOracle.charts(chart.workbookPath(), "Chart");
    XlsxParityOracle.DrawingSheetSnapshot directDrawing =
        XlsxParityOracle.drawingSheet(chart.workbookPath(), "Chart");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            chart.workbookPath(),
            inspect(
                "charts", new ChartSelector.AllOnSheet("Chart"), new InspectionQuery.GetCharts()),
            inspect(
                "drawing",
                new DrawingObjectSelector.AllOnSheet("Chart"),
                new InspectionQuery.GetDrawingObjects()));
    InspectionResult.ChartsResult charts =
        XlsxParityGridGrind.read(success, "charts", InspectionResult.ChartsResult.class);
    InspectionResult.DrawingObjectsResult drawing =
        XlsxParityGridGrind.read(success, "drawing", InspectionResult.DrawingObjectsResult.class);
    Path replacementPath = context.derivedWorkbook("chart-combo-replaced");
    XlsxParityGridGrind.mutateWorkbook(
        chart.workbookPath(),
        replacementPath,
        List.of(
            mutate(
                new SheetSelector.ByName("Chart"),
                new MutationAction.SetChart(
                    new ChartInput(
                        "ComboChart",
                        twoCellAnchorInput(
                            4, 1, 11, 16, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                        chartTitle("Roadmap"),
                        new ChartInput.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                        ExcelChartDisplayBlanksAs.SPAN,
                        false,
                        List.of(
                            new ChartInput.Bar(
                                true,
                                ExcelChartBarDirection.COLUMN,
                                null,
                                null,
                                null,
                                null,
                                List.of(
                                    new ChartInput.Series(
                                        new ChartInput.Title.Formula("B1"),
                                        new ChartInput.DataSource.Reference("A2:A4"),
                                        new ChartInput.DataSource.Reference("B2:B4"),
                                        null,
                                        null,
                                        null,
                                        null)))))))));
    if (directCharts.size() != 1
        || directCharts.getFirst().plots().size() != 2
        || directDrawing.objects().size() != 1
        || !(directDrawing.objects().getFirst()
            instanceof XlsxParityOracle.ChartDrawingObjectSnapshot directDrawingChart)
        || charts.charts().size() != 1
        || drawing.drawingObjects().size() != 1
        || !(drawing.drawingObjects().getFirst()
            instanceof DrawingObjectReport.Chart reportChart)) {
      return fail(
          "Combo chart corpus workbook did not surface the expected multi-plot chart shape.");
    }
    InspectionResult.ChartsResult replacedCharts =
        XlsxParityGridGrind.read(
            XlsxParityGridGrind.readWorkbook(
                replacementPath,
                inspect(
                    "charts",
                    new ChartSelector.AllOnSheet("Chart"),
                    new InspectionQuery.GetCharts())),
            "charts",
            InspectionResult.ChartsResult.class);
    boolean matches =
        charts.charts().equals(directCharts)
            && chartDrawingMatches(reportChart, directDrawingChart)
            && plotTypeTokens(directCharts.getFirst()).equals(List.of("BAR", "LINE"))
            && replacedCharts.charts().size() == 1
            && hasSinglePlot(replacedCharts.charts().getFirst(), ChartReport.Bar.class);
    return matches
        ? pass(
            "Combo charts are surfaced truthfully, preserved losslessly, and can be authoritatively replaced.")
        : fail("Combo-chart handling diverged from the direct POI oracle or replacement contract.");
  }

  static ProbeResult probePivotPreservation(ProbeContext context) {
    List<String> mismatches = new ArrayList<>();
    for (String scenarioId :
        List.of(XlsxParityScenarios.PIVOT, XlsxParityScenarios.PIVOT_AUTHORING)) {
      XlsxParityScenarios.MaterializedScenario pivot =
          context.copiedScenario(scenarioId, "pivot-preservation-" + scenarioId);
      List<PivotTableReport> before = XlsxParityOracle.pivotTables(pivot.workbookPath());
      Path outputPath = context.derivedWorkbook("pivot-preserved-" + scenarioId.replace('-', '_'));
      XlsxParityGridGrind.mutateWorkbook(
          pivot.workbookPath(),
          outputPath,
          List.of(
              mutate(new SheetSelector.ByName("Scratch"), new MutationAction.EnsureSheet()),
              mutate(
                  new CellSelector.ByAddress("Scratch", "A1"),
                  new MutationAction.SetCell(text("Touch")))));
      List<PivotTableReport> after = XlsxParityOracle.pivotTables(outputPath);
      if (!before.equals(after)) {
        mismatches.add(scenarioId);
      }
    }
    return mismatches.isEmpty()
        ? pass(
            "Phase 7 pivot corpus workbooks preserve pivot parts and sources across unrelated edits.")
        : fail("Unrelated GridGrind edits changed pivot state for scenarios: " + mismatches);
  }

  static ProbeResult probePivotReadback(ProbeContext context) {
    if (!XlsxParityGridGrind.hasInspectionQueryType("GET_PIVOT_TABLES")
        || !XlsxParityGridGrind.hasInspectionQueryType("ANALYZE_PIVOT_TABLE_HEALTH")) {
      return fail("GridGrind is missing the Phase 7 pivot read contract.");
    }
    XlsxParityScenarios.MaterializedScenario pivot = context.scenario(XlsxParityScenarios.PIVOT);
    List<PivotTableReport> directPivots = XlsxParityOracle.pivotTables(pivot.workbookPath());
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            pivot.workbookPath(),
            inspect("pivots", new PivotTableSelector.All(), new InspectionQuery.GetPivotTables()),
            inspect(
                "health",
                new PivotTableSelector.All(),
                new InspectionQuery.AnalyzePivotTableHealth()));
    InspectionResult.PivotTablesResult pivots =
        XlsxParityGridGrind.read(success, "pivots", InspectionResult.PivotTablesResult.class);
    InspectionResult.PivotTableHealthResult health =
        XlsxParityGridGrind.read(success, "health", InspectionResult.PivotTableHealthResult.class);
    if (directPivots.size() != 1
        || !(directPivots.getFirst() instanceof PivotTableReport.Supported supported)
        || !(supported.source() instanceof PivotTableReport.Source.Range)
        || !supported.rowLabels().equals(List.of(new PivotTableReport.Field(0, "Region")))) {
      return fail(
          "Direct POI pivot corpus workbook is not in the expected supported single-pivot shape.");
    }
    boolean matches = pivots.pivotTables().equals(directPivots) && pivotHealthClean(health, 1);
    return matches
        ? pass("GridGrind reads supported POI-authored pivot tables with direct POI agreement.")
        : fail("GridGrind pivot readback diverged from the direct POI oracle.");
  }

  static ProbeResult probePivotAuthoring(ProbeContext context) {
    if (!XlsxParityGridGrind.hasMutationActionType("SET_PIVOT_TABLE")
        || !XlsxParityGridGrind.hasInspectionQueryType("GET_PIVOT_TABLES")
        || !XlsxParityGridGrind.hasInspectionQueryType("ANALYZE_PIVOT_TABLE_HEALTH")) {
      return fail("GridGrind is missing the Phase 7 pivot authoring contract.");
    }
    XlsxParityScenarios.MaterializedScenario pivot =
        context.scenario(XlsxParityScenarios.PIVOT_AUTHORING);
    List<PivotTableReport> directPivots = XlsxParityOracle.pivotTables(pivot.workbookPath());
    GridGrindResponse.Success success =
        XlsxParityGridGrind.readWorkbook(
            pivot.workbookPath(),
            inspect("pivots", new PivotTableSelector.All(), new InspectionQuery.GetPivotTables()),
            inspect(
                "health",
                new PivotTableSelector.All(),
                new InspectionQuery.AnalyzePivotTableHealth()));
    InspectionResult.PivotTablesResult pivots =
        XlsxParityGridGrind.read(success, "pivots", InspectionResult.PivotTablesResult.class);
    InspectionResult.PivotTableHealthResult health =
        XlsxParityGridGrind.read(success, "health", InspectionResult.PivotTableHealthResult.class);
    if (directPivots.size() != 3
        || supportedPivotByName(directPivots, "Sales Pivot") == null
        || !(supportedPivotByName(directPivots, "Named Pivot").source()
            instanceof PivotTableReport.Source.NamedRange)
        || !(supportedPivotByName(directPivots, "Table Pivot").source()
            instanceof PivotTableReport.Source.Table)) {
      return fail(
          "GridGrind-authored pivot corpus workbook did not retain the expected range, named-range, and table-backed pivot shapes.");
    }
    boolean matches =
        pivots.pivotTables().equals(directPivots) && pivotHealthClean(health, directPivots.size());
    return matches
        ? pass(
            "GridGrind authors range-, named-range-, and table-backed pivot tables and rereads them with direct POI agreement.")
        : fail("GridGrind-authored pivot workbook diverged from the direct POI oracle.");
  }

  static ProbeResult probePivotMutation(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario pivot =
        context.copiedScenario(XlsxParityScenarios.PIVOT_AUTHORING, "pivot-mutation-source");
    Path outputPath = context.derivedWorkbook("pivot-mutated");
    GridGrindResponse.Success success =
        XlsxParityGridGrind.mutateWorkbook(
            pivot.workbookPath(),
            outputPath,
            List.of(
                mutate(
                    new MutationAction.SetPivotTable(
                        new PivotTableInput(
                            "Sales Pivot",
                            "RangeReport",
                            new PivotTableInput.Source.Range("Data", "A1:D5"),
                            new PivotTableInput.Anchor("D6"),
                            List.of("Stage"),
                            List.of("Region"),
                            List.of(),
                            List.of(
                                new PivotTableInput.DataField(
                                    "Amount",
                                    dev.erst.gridgrind.excel.foundation
                                        .ExcelPivotDataConsolidateFunction.SUM,
                                    "Total Amount",
                                    "#,##0.00"))))),
                mutate(
                    new PivotTableSelector.ByNameOnSheet("Table Pivot", "TableReport"),
                    new MutationAction.DeletePivotTable())),
            inspect("pivots", new PivotTableSelector.All(), new InspectionQuery.GetPivotTables()),
            inspect(
                "health",
                new PivotTableSelector.All(),
                new InspectionQuery.AnalyzePivotTableHealth()));
    List<PivotTableReport> directPivots = XlsxParityOracle.pivotTables(outputPath);
    InspectionResult.PivotTablesResult pivots =
        XlsxParityGridGrind.read(success, "pivots", InspectionResult.PivotTablesResult.class);
    InspectionResult.PivotTableHealthResult health =
        XlsxParityGridGrind.read(success, "health", InspectionResult.PivotTableHealthResult.class);
    PivotTableReport.Supported salesPivot = supportedPivotByName(directPivots, "Sales Pivot");
    if (directPivots.size() != 2
        || salesPivot == null
        || !"D6".equals(salesPivot.anchor().topLeftAddress())
        || !salesPivot.rowLabels().equals(List.of(new PivotTableReport.Field(1, "Stage")))
        || !salesPivot.columnLabels().equals(List.of(new PivotTableReport.Field(0, "Region")))
        || supportedPivotByName(directPivots, "Table Pivot") != null) {
      return fail(
          "Pivot mutation parity did not produce the expected updated supported pivot set.");
    }
    boolean matches =
        pivots.pivotTables().equals(directPivots) && pivotHealthClean(health, directPivots.size());
    return matches
        ? pass("GridGrind mutates and deletes supported pivot tables with direct POI agreement.")
        : fail("GridGrind pivot mutation diverged from the direct POI oracle.");
  }

  private static boolean pivotHealthClean(
      InspectionResult.PivotTableHealthResult health, int expectedPivotCount) {
    return health.analysis().checkedPivotTableCount() == expectedPivotCount
        && health.analysis().summary().totalCount() == 0
        && health.analysis().findings().isEmpty();
  }

  private static PivotTableReport.Supported supportedPivotByName(
      List<PivotTableReport> pivots, String name) {
    return pivots.stream()
        .filter(PivotTableReport.Supported.class::isInstance)
        .map(PivotTableReport.Supported.class::cast)
        .filter(pivot -> pivot.name().equals(name))
        .findFirst()
        .orElse(null);
  }
}
