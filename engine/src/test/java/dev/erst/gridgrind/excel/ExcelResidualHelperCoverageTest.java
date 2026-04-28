package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import com.microsoft.schemas.office.excel.CTClientData;
import com.microsoft.schemas.vml.CTShape;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartMarkerStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFArea3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBar3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLine3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPie3DChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt;
import org.openxmlformats.schemas.drawingml.x2006.chart.STMarkerStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTLegacyDrawing;

/** Residual helper and private-seam coverage for rebuilt XLSX engine surfaces. */
class ExcelResidualHelperCoverageTest {
  @Test
  @SuppressWarnings("PMD.NcssCount")
  void chartHelperSeamsCoverResidualBranches() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Charts");
      seedChartData(sheet);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();

      XSSFChart emptyChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 1, 7, 10));
      assertFalse(ExcelChartSnapshotSupport.barVaryColors(emptyChart));
      assertFalse(ExcelChartSnapshotSupport.lineVaryColors(emptyChart));
      XSSFChart varyingBarChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 11, 7, 20));
      varyingBarChart.getCTChart().getPlotArea().addNewBarChart().addNewVaryColors().setVal(true);
      assertTrue(ExcelChartSnapshotSupport.barVaryColors(varyingBarChart));
      XSSFChart varyingLineChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 8, 11, 14, 20));
      varyingLineChart.getCTChart().getPlotArea().addNewLineChart().addNewVaryColors().setVal(true);
      assertTrue(ExcelChartSnapshotSupport.lineVaryColors(varyingLineChart));

      List<?> emptyPlots = ExcelChartPlotSnapshotSupport.snapshotPlots(emptyChart, List.of());
      assertInstanceOf(ExcelChartSnapshot.Unsupported.class, emptyPlots.getFirst());

      XDDFBarChartData unplottedBarData =
          (XDDFBarChartData)
              emptyChart.createData(
                  ChartTypes.BAR,
                  emptyChart.createCategoryAxis(AxisPosition.BOTTOM),
                  emptyChart.createValueAxis(AxisPosition.LEFT));
      List<?> unsupportedPlots =
          ExcelChartPlotSnapshotSupport.snapshotPlots(
              emptyChart, List.of(unplottedBarData, new UnsupportedChartData()));
      assertEquals(2, unsupportedPlots.size());
      assertTrue(
          unsupportedPlots.stream().allMatch(ExcelChartSnapshot.Unsupported.class::isInstance));

      XSSFChart lineChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 8, 1, 14, 10));
      XDDFLineChartData lineData =
          (XDDFLineChartData)
              lineChart.createData(
                  ChartTypes.LINE,
                  lineChart.createCategoryAxis(AxisPosition.BOTTOM),
                  lineChart.createValueAxis(AxisPosition.LEFT));
      XDDFLineChartData.Series lineSeries =
          (XDDFLineChartData.Series)
              lineData.addSeries(
                  XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Feb"}),
                  XDDFDataSourcesFactory.fromArray(new Double[] {10d, 12d}));
      for (ExcelChartMarkerStyle markerStyle : ExcelChartMarkerStyle.values()) {
        lineSeries.setMarkerStyle(ExcelChartPoiBridge.toPoiMarkerStyle(markerStyle));
        assertEquals(markerStyle, invokeMarkerStyle(lineSeries.getCTLineSer().getMarker()));
      }
      lineSeries.setMarkerSize((short) 11);
      assertEquals(
          Short.valueOf((short) 11), invokeMarkerSize(lineSeries.getCTLineSer().getMarker()));
      assertNull(invokeMarkerStyle(CTMarker.Factory.newInstance()));
      CTMarker autoMarker = CTMarker.Factory.newInstance();
      autoMarker.addNewSymbol().setVal(STMarkerStyle.AUTO);
      assertNull(invokeMarkerStyle(autoMarker));
      assertNull(invokeMarkerSize(CTMarker.Factory.newInstance()));
      CTMarker unsupportedMarker =
          CTMarker.Factory.parse(
              "<c:marker xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart'>"
                  + "<c:symbol val='hexagon'/></c:marker>");
      assertNull(invokeMarkerStyle(unsupportedMarker));

      ExcelChartAxisRegistry registry =
          new ExcelChartAxisRegistry(
              drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 15, 1, 22, 10)));
      assertThrows(
          IllegalArgumentException.class,
          () -> registry.categoryValueAxes(List.of(axis(ExcelChartAxisKind.SERIES))));
      assertThrows(
          IllegalArgumentException.class,
          () -> registry.categoryValueAxes(List.of(axis(ExcelChartAxisKind.CATEGORY))));
      assertThrows(
          IllegalArgumentException.class,
          () -> registry.categoryValueAxes(List.of(axis(ExcelChartAxisKind.VALUE))));
      ExcelChartAxisRegistry.CategoryValueAxes dateAxes =
          registry.categoryValueAxes(
              List.of(axis(ExcelChartAxisKind.DATE), axis(ExcelChartAxisKind.VALUE)));
      assertNotNull(dateAxes.categoryAxis());
      assertNotNull(dateAxes.valueAxis());
      ExcelChartAxisRegistry.SurfaceAxes surfaceAxes =
          registry.surfaceAxes(
              List.of(
                  axis(ExcelChartAxisKind.DATE),
                  axis(ExcelChartAxisKind.VALUE),
                  axis(ExcelChartAxisKind.SERIES)));
      assertNotNull(surfaceAxes.categoryAxis());
      assertNotNull(surfaceAxes.valueAxis());
      assertNotNull(surfaceAxes.seriesAxis());
      assertThrows(
          IllegalArgumentException.class,
          () ->
              registry.surfaceAxes(
                  List.of(axis(ExcelChartAxisKind.DATE), axis(ExcelChartAxisKind.VALUE))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              registry.surfaceAxes(
                  List.of(axis(ExcelChartAxisKind.DATE), axis(ExcelChartAxisKind.SERIES))));

      Name localNumbers = workbook.createName();
      localNumbers.setNameName("LocalNums");
      localNumbers.setSheetIndex(workbook.getSheetIndex(sheet));
      localNumbers.setRefersToFormula("B2:B3");
      Name workbookScopedNumbers = workbook.createName();
      workbookScopedNumbers.setNameName("WorkbookNums");
      workbookScopedNumbers.setRefersToFormula("B2:B3");
      Name sparseNumbers = workbook.createName();
      sparseNumbers.setNameName("SparseNums");
      sparseNumbers.setRefersToFormula("B20:B21");
      XDDFValueAxis scratchAxis = lineChart.createValueAxis(AxisPosition.RIGHT);
      assertEquals(
          2,
          ExcelChartSourceSupport.toValueDataSource(
                  sheet, new ExcelChartDefinition.DataSource.Reference("LocalNums"))
              .getPointCount());
      assertTrue(ExcelChartSourceSupport.resolveChartSource(sheet, "LocalNums").numeric());
      assertTrue(ExcelChartSourceSupport.resolveChartSource(sheet, "WorkbookNums").numeric());
      ResolvedChartSource sparseResolved =
          ExcelChartSourceSupport.resolveChartSource(sheet, "SparseNums");
      assertFalse(sparseResolved.numeric());
      assertEquals(List.of("", ""), sparseResolved.stringValues());
      assertEquals(
          2,
          ExcelChartSourceSupport.toValueDataSource(
                  sheet, new ExcelChartDefinition.DataSource.NumericLiteral(List.of(3d, 4d)))
              .getPointCount());
      assertThrows(
          IllegalArgumentException.class,
          () ->
              ExcelChartSourceSupport.toValueDataSource(
                  sheet, new ExcelChartDefinition.DataSource.StringLiteral(List.of("Jan"))));

      CTUnsignedInt knownAxisId = CTUnsignedInt.Factory.newInstance();
      knownAxisId.setVal(scratchAxis.getId());
      CTUnsignedInt missingAxisId = CTUnsignedInt.Factory.newInstance();
      missingAxisId.setVal(Long.MAX_VALUE);
      List<?> snapshotAxes =
          ExcelChartPlotSnapshotSupport.snapshotAxes(
              java.util.Map.of(scratchAxis.getId(), (XDDFChartAxis) scratchAxis),
              List.of(knownAxisId, missingAxisId));
      assertEquals(1, snapshotAxes.size());

      assertThrows(
          IllegalStateException.class,
          () -> ExcelChartPlotSnapshotSupport.snapshotDataSource(sheet, null));
      assertInstanceOf(
          ExcelChartSnapshot.DataSource.NumericLiteral.class,
          ExcelChartPlotSnapshotSupport.snapshotDataSource(
              sheet, XDDFDataSourcesFactory.fromArray(new Double[] {1d, 2d})));

      ExcelChartMutationSupport.createChart(
          sheet,
          new ExcelChartDefinition(
              "Bar3DNoShape",
              ExcelChartTestSupport.anchor(23, 1, 29, 10),
              new ExcelChartDefinition.Title.Text("Bar 3D"),
              new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.RIGHT),
              ExcelChartDisplayBlanksAs.GAP,
              true,
              List.of(
                  new ExcelChartDefinition.Bar3D(
                      false,
                      ExcelChartBarDirection.COLUMN,
                      ExcelChartBarGrouping.CLUSTERED,
                      100,
                      50,
                      null,
                      List.of(axis(ExcelChartAxisKind.CATEGORY), axis(ExcelChartAxisKind.VALUE)),
                      List.of(
                          new ExcelChartDefinition.Series(
                              null,
                              new ExcelChartDefinition.DataSource.Reference("A2:A4"),
                              new ExcelChartDefinition.DataSource.Reference("B2:B4")))))));

      XSSFChart mutationChart =
          drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 30, 1, 36, 10));
      applySeriesOptionBranches(mutationChart);
      applySeriesTitleAndChartTitleBranches(mutationChart, sheet);
      snapshotPlotDefaultsForUnsetGroupingAndShape(
          drawing,
          XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Feb"}),
          XDDFDataSourcesFactory.fromArray(new Double[] {10d, 12d}));
    }
  }

  @Test
  @SuppressWarnings("PMD.NcssCount")
  void drawingSignatureAndSheetHelpersCoverResidualBranches() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      ExcelDrawingController controller = new ExcelDrawingController();
      ExcelSignatureLineController signatureLineController = new ExcelSignatureLineController();
      ExcelDrawingAnchor.TwoCell anchor = ExcelChartTestSupport.anchor(1, 1, 4, 6);

      controller.setSignatureLine(sheet, signatureDefinition("OpsSignature", anchor));
      assertFalse(signatureLineController.hasNamedSignatureLine(sheet, "Missing"));
      CTShape signatureShape = signatureShapes(sheet.getVMLDrawing(false)).getFirst();
      CTClientData clientData = signatureShape.getClientDataArray(0);
      String originalAnchor = clientData.getAnchorArray(0);
      ExcelDrawingAnchor.TwoCell movedAnchor = ExcelChartTestSupport.anchor(2, 2, 5, 7);
      assertTrue(signatureLineController.updateAnchorIfPresent(sheet, "OpsSignature", movedAnchor));
      assertEquals(1, clientData.sizeOfAnchorArray());
      assertNotEquals(originalAnchor, clientData.getAnchorArray(0));

      org.apache.poi.openxml4j.opc.PackagePart previewPart =
          ExcelDrawingBinarySupport.relatedInternalPart(
              sheet.getVMLDrawing(false).getPackagePart(),
              signatureShape.getImagedataArray(0).getRelid());
      assertTrue(ExcelDrawingBinarySupport.imagePartUsed(workbook, previewPart.getPartName()));
      sheet.createDrawingPatriarch();
      assertTrue(ExcelDrawingBinarySupport.imagePartUsed(workbook, previewPart.getPartName()));

      signatureShape.removeImagedata(0);
      ExcelSignatureLineSnapshot snapshot =
          signatureLineController.signatureLines(sheet).getFirst();
      assertNull(snapshot.previewFormat());
      assertNull(snapshot.previewContentType());
      assertTrue(signatureLineController.deleteIfPresent(sheet, "OpsSignature"));
      assertFalse(signatureLineController.deleteIfPresent(sheet, "OpsSignature"));
      assertFalse(signatureLineController.hasNamedSignatureLine(sheet, "OpsSignature"));
      assertThrows(
          IllegalArgumentException.class,
          () -> signatureLineController.deleteIfPresent(sheet, " "));

      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFSimpleShape groupChild =
          drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 1, 8, 4, 11));
      groupChild.getCTShape().getNvSpPr().getCNvPr().setName("Mover");
      controller.setDrawingObjectAnchor(sheet, "Mover", ExcelChartTestSupport.anchor(2, 8, 5, 11));
      controller.deleteDrawingObject(sheet, "Mover");
      assertThrows(
          DrawingObjectNotFoundException.class,
          () -> controller.deleteDrawingObject(sheet, "Missing"));

      var group = drawing.createGroup(drawing.createAnchor(0, 0, 0, 0, 6, 8, 9, 12));
      group.getCTGroupShape().getNvGrpSpPr().getCNvPr().setName("Grouped");
      IllegalArgumentException readOnlyGroup =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setDrawingObjectAnchor(
                      sheet, "Grouped", ExcelChartTestSupport.anchor(10, 8, 12, 12)));
      assertTrue(readOnlyGroup.getMessage().contains("read-only"));

      assertInstanceOf(
          DrawingObjectNotFoundException.class,
          assertThrows(
              DrawingObjectNotFoundException.class,
              () -> controller.requiredLocatedShape(sheet, "Missing")));

      XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 1, 13, 7, 20));
      XSSFGraphicFrame graphicFrame = chart.getGraphicFrame();
      ExcelDrawingController.LocatedShape locatedChart =
          new ExcelDrawingController.LocatedShape(
              drawing,
              graphicFrame,
              ExcelDrawingAnchorSupport.shapeXml(graphicFrame),
              ExcelDrawingAnchorSupport.parentAnchor(
                  ExcelDrawingAnchorSupport.shapeXml(graphicFrame)));
      IllegalStateException chartRemovalFailure =
          assertThrows(
              IllegalStateException.class,
              () ->
                  ExcelDrawingRemovalSupport.deleteLocatedShape(
                      sheet, locatedChart, (left, right) -> false));
      assertTrue(chartRemovalFailure.getMessage().contains("Failed to remove chart relation"));

      ExcelDrawingObjectSnapshot.Shape graphicFrameSnapshot =
          (ExcelDrawingObjectSnapshot.Shape)
              ExcelDrawingSnapshotSupport.snapshotGraphicFrame(graphicFrame);
      assertEquals(ExcelDrawingShapeKind.GRAPHIC_FRAME, graphicFrameSnapshot.kind());

      XSSFSheet chartSheet = workbook.createSheet("ChartTarget");
      seedChartData(chartSheet);
      IllegalArgumentException invalidSurfaceFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  controller.setChart(
                      chartSheet,
                      new ExcelChartDefinition(
                          "BrokenSurface",
                          ExcelChartTestSupport.anchor(1, 1, 7, 10),
                          new ExcelChartDefinition.Title.Text("Broken"),
                          new ExcelChartDefinition.Legend.Visible(ExcelChartLegendPosition.RIGHT),
                          ExcelChartDisplayBlanksAs.GAP,
                          true,
                          List.of(
                              new ExcelChartDefinition.Surface(
                                  true,
                                  true,
                                  List.of(
                                      axis(ExcelChartAxisKind.CATEGORY),
                                      axis(ExcelChartAxisKind.VALUE)),
                                  List.of(
                                      new ExcelChartDefinition.Series(
                                          null,
                                          new ExcelChartDefinition.DataSource.Reference("A2:A4"),
                                          new ExcelChartDefinition.DataSource.Reference(
                                              "B2:B4"))))))));
      assertTrue(invalidSurfaceFailure.getMessage().contains("Surface plots must declare"));
      assertTrue(chartSheet.createDrawingPatriarch().getCharts().isEmpty());

      XSSFSheet annotationSheet = workbook.createSheet("Annotations");
      ExcelSheetAnnotationSupport annotationSupport =
          new ExcelSheetAnnotationSupport(annotationSheet, new ExcelDrawingController());
      annotationSupport.setComment("A1", new ExcelComment("Review", "GridGrind", false));
      String vmlRelationId =
          ExcelSheetAnnotationSupport.vmlDrawingRelationId(annotationSheet).orElseThrow();
      assertEquals(
          java.util.Optional.of(vmlRelationId),
          ExcelSheetAnnotationSupport.legacyDrawingRelationId(annotationSheet));
      CTLegacyDrawing legacyDrawing = annotationSheet.getCTWorksheet().getLegacyDrawing();
      legacyDrawing.setId("rIdMissing");
      ExcelSheetAnnotationSupport.repairBrokenLegacyDrawingReference(annotationSheet);
      assertFalse(annotationSheet.getCTWorksheet().isSetLegacyDrawing());
      ExcelSheetAnnotationSupport.ensureLegacyDrawingReference(annotationSheet);
      assertEquals(
          java.util.Optional.of(vmlRelationId),
          ExcelSheetAnnotationSupport.legacyDrawingRelationId(annotationSheet));
      annotationSupport.clearComment("A1");
      annotationSheet.createRow(5).createCell(5);
      ExcelSheetAnnotationSupport.clearCellComment(annotationSheet.getRow(5).getCell(5));
      try (HSSFWorkbook legacyWorkbook = new HSSFWorkbook()) {
        ExcelSheetAnnotationSupport.clearCellComment(
            legacyWorkbook.createSheet("Legacy").createRow(0).createCell(0));
      }
      XSSFSheet blankAnnotationSheet = workbook.createSheet("Blank");
      assertTrue(ExcelSheetAnnotationSupport.vmlDrawingRelationId(blankAnnotationSheet).isEmpty());
      assertTrue(
          ExcelSheetAnnotationSupport.legacyDrawingRelationId(blankAnnotationSheet).isEmpty());
      ExcelSheetAnnotationSupport.ensureLegacyDrawingReference(blankAnnotationSheet);
      assertFalse(blankAnnotationSheet.getCTWorksheet().isSetLegacyDrawing());
      ExcelSheetAnnotationSupport.removeCommentFromTable(
          blankAnnotationSheet, new CellAddress("A1"));
      ExcelSheetAnnotationSupport.removeCommentShapeIfPresent(
          annotationSheet, new CellAddress("A1"));

      ExcelSheetPresentationController.clearTabColor(workbook.createSheet("NoSheetPr"));
      XSSFSheet colorSheet = workbook.createSheet("Color");
      colorSheet.setTabColor(ExcelColorSupport.toXssfColor(workbook, ExcelColor.rgb("#FF0000")));
      ExcelSheetPresentationController.clearTabColor(colorSheet);
      assertNull(colorSheet.getTabColor());

      XSSFSheet paneSheet = workbook.createSheet("Pane");
      ExcelSheetViewSupport.setPane(paneSheet, new ExcelSheetPane.None());
      assertInstanceOf(ExcelSheetPane.None.class, ExcelSheetViewSupport.pane(paneSheet));
    }
  }

  @Test
  void workbookCommandsArrayFormulaAndTempHelpersCoverResidualBranches() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setRange(
          "B2:C4",
          List.of(
              List.of(ExcelCellValue.number(2d), ExcelCellValue.number(3d)),
              List.of(ExcelCellValue.number(4d), ExcelCellValue.number(5d)),
              List.of(ExcelCellValue.number(6d), ExcelCellValue.number(7d))));
      sheet.setCell("A1", ExcelCellValue.formula("1+1"));
      assertEquals(List.of(), sheet.arrayFormulas());

      WorkbookCommandExecutor.applyCellValueCommand(
          workbook,
          new WorkbookCommand.SetArrayFormula(
              "Ops", "D2:D4", new ExcelArrayFormulaDefinition("B2:B4*C2:C4")));
      assertEquals(1, sheet.arrayFormulas().size());

      WorkbookCommandExecutor.applyCellValueCommand(
          workbook, new WorkbookCommand.ClearArrayFormula("Ops", "D2"));
      assertEquals(List.of(), sheet.arrayFormulas());

      assertThrows(
          InvalidFormulaException.class,
          () -> sheet.setArrayFormula("E2:E4", new ExcelArrayFormulaDefinition("SUM(")));
      assertThrows(CellNotFoundException.class, () -> sheet.clearArrayFormula("Z99"));
      sheet.xssfSheet().createRow(0);
      assertThrows(CellNotFoundException.class, () -> sheet.clearArrayFormula("B1"));
      assertThrows(InvalidCellAddressException.class, () -> sheet.clearArrayFormula("XFE1048577"));
      assertThrows(
          InvalidCellAddressException.class,
          () ->
              ExcelSheetCellMutationSupport.requireValidCellReference(
                  "A0", new CellReference(-1, 0)));
      assertThrows(
          InvalidCellAddressException.class,
          () ->
              ExcelSheetCellMutationSupport.requireValidCellReference(
                  "A0", new CellReference(0, -1)));
      assertThrows(
          InvalidCellAddressException.class,
          () ->
              ExcelSheetCellMutationSupport.requireValidCellReference(
                  "A0",
                  new CellReference(0, SpreadsheetVersion.EXCEL2007.getLastColumnIndex() + 1)));
      assertThrows(
          InvalidCellAddressException.class,
          () ->
              ExcelSheetCellMutationSupport.requireValidCellReference(
                  "A1048577",
                  new CellReference(SpreadsheetVersion.EXCEL2007.getLastRowIndex() + 1, 0)));

      WorkbookCommandExecutor.applyWorkbookMetadataCommand(
          workbook,
          new WorkbookCommand.SetSignatureLine(
              "Ops",
              signatureDefinition("OpsSignature", ExcelChartTestSupport.anchor(1, 8, 4, 12))));
      WorkbookCommandExecutor.applyWorkbookMetadataCommand(
          workbook,
          new WorkbookCommand.SetNamedRange(
              new ExcelNamedRangeDefinition(
                  "BudgetTotal",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  new ExcelNamedRangeTarget("Ops", "B2"))));
      WorkbookCommandExecutor.applyWorkbookMetadataCommand(
          workbook,
          new WorkbookCommand.DeleteNamedRange(
              "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope()));
      WorkbookCommandExecutor.applyWorkbookMetadataCommand(
          workbook,
          new WorkbookCommand.SetShape(
              "Ops",
              new ExcelShapeDefinition(
                  "Mover",
                  ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                  ExcelChartTestSupport.anchor(5, 8, 8, 12),
                  "rect",
                  "text")));
      WorkbookCommandExecutor.applyWorkbookMetadataCommand(
          workbook,
          new WorkbookCommand.SetDrawingObjectAnchor(
              "Ops", "Mover", ExcelChartTestSupport.anchor(9, 8, 12, 12)));
      WorkbookCommandExecutor.applyWorkbookMetadataCommand(
          workbook, new WorkbookCommand.DeleteDrawingObject("Ops", "Mover"));
    }

    try (ExcelWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      WorkbookCommandExecutor.applyWorkbookMetadataCommand(
          workbook,
          new WorkbookCommand.ImportCustomXmlMapping(
              new ExcelCustomXmlImportDefinition(
                  new ExcelCustomXmlMappingLocator(1L, "CORSO_mapping"),
                  "<CORSO><NOME>Ops</NOME><DOCENTE>Agent</DOCENTE><TUTOR>Grid</TUTOR><CDL>QA</CDL>"
                      + "<DURATA>4</DURATA><ARGOMENTO>Coverage</ARGOMENTO>"
                      + "<PROGETTO>Parity</PROGETTO><CREDITI>8</CREDITI></CORSO>")));
      assertEquals(
          "Ops",
          workbook
              .sheet("Foglio1")
              .window("A1", 1, 1)
              .rows()
              .getFirst()
              .cells()
              .getFirst()
              .displayValue());
    }

    String originalSystemTemp = System.getProperty("java.io.tmpdir");
    String originalUserHome = System.getProperty("user.home");
    System.clearProperty("java.io.tmpdir");
    try {
      assertNull(ExcelTempFiles.systemTempRoot());
      Path fallbackOnlyRoot = java.nio.file.Files.createTempDirectory("gridgrind-temp-home-only-");
      System.setProperty("user.home", fallbackOnlyRoot.toString());
      Path tempFile = ExcelTempFiles.createManagedTempFile("gridgrind-home-only-", ".tmp");
      assertTrue(tempFile.startsWith(fallbackOnlyRoot.resolve(".gridgrind").resolve("tmp")));
      java.nio.file.Files.deleteIfExists(tempFile);
      try (java.util.stream.Stream<Path> paths =
          java.nio.file.Files.walk(fallbackOnlyRoot.resolve(".gridgrind"))) {
        paths
            .sorted(java.util.Comparator.reverseOrder())
            .forEach(
                path -> {
                  try {
                    java.nio.file.Files.deleteIfExists(path);
                  } catch (java.io.IOException ignored) {
                    // best-effort cleanup for test-owned temp roots
                  }
                });
      }
      java.nio.file.Files.deleteIfExists(fallbackOnlyRoot);
    } finally {
      restoreProperty("java.io.tmpdir", originalSystemTemp);
      restoreProperty("user.home", originalUserHome);
    }

    System.clearProperty("user.home");
    try {
      assertNull(ExcelTempFiles.userHomeFallbackRoot());
      Path tempFile = ExcelTempFiles.createManagedTempFile("gridgrind-system-only-", ".tmp");
      assertEquals("gridgrind", tempFile.getParent().getFileName().toString());
      java.nio.file.Files.deleteIfExists(tempFile);
    } finally {
      restoreProperty("user.home", originalUserHome);
    }
  }

  @Test
  void customXmlAndSheetCopyResidualHelpersCoverBranches() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Source");
      workbook.getOrCreateSheet("Target");
      workbook.sheet("Target").setCell("A1", ExcelCellValue.formula("1+1"));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "LocalName",
              new ExcelNamedRangeScope.SheetScope("Target"),
              new ExcelNamedRangeTarget("Target", "A1")));
      workbook.setNamedRange(
          new ExcelNamedRangeDefinition(
              "WorkbookName",
              new ExcelNamedRangeScope.WorkbookScope(),
              new ExcelNamedRangeTarget("Target", "A1")));

      ExcelSheetCopyController.deleteLocalNamedRanges(workbook, "Target");
      assertEquals(
          List.of("WorkbookName"),
          workbook.namedRanges().stream().map(ExcelNamedRangeSnapshot::name).toList());
      assertTrue(ExcelSheetCopyController.mayReferenceCopiedSheet("Source!A1", "Source"));
      assertTrue(
          ExcelSheetCopyController.mayReferenceCopiedSheet("'Source Name'!A1", "Source Name"));
      assertTrue(
          ExcelSheetCopyController.mayReferenceCopiedSheet("SUM(Source:Target!A1)", "Source"));
      assertTrue(
          ExcelSheetCopyController.mayReferenceCopiedSheet(
              "SUM('Source Name':'Target'!A1)", "Source Name"));
      assertFalse(ExcelSheetCopyController.mayReferenceCopiedSheet("Other!A1", "Source"));
      workbook
          .sheet("Target")
          .xssfSheet()
          .getRow(0)
          .createCell(1)
          .setCellFormula("\"Source:note\"");
      ExcelSheetCopyController.retargetCopiedSheetFormulas(
          workbook, "Source", "Target Copy", workbook.sheet("Target").xssfSheet());
      assertEquals(
          "1+1", workbook.sheet("Target").xssfSheet().getRow(0).getCell(0).getCellFormula());
      assertEquals(
          "\"Source:note\"",
          workbook.sheet("Target").xssfSheet().getRow(0).getCell(1).getCellFormula());
    }

    try (ExcelWorkbook workbook = CustomXmlWorkbookSamples.openSimpleCustomXmlWorkbook()) {
      var mapping = workbook.xssfWorkbook().getCustomXMLMappings().iterator().next();
      assertNotNull(ExcelCustomXmlController.schemaXml(mapping));
      assertNotNull(
          ExcelCustomXmlController.schemaXml(
              mapping, org.apache.poi.util.XMLHelper.newTransformer()));
      assertThrows(
          NullPointerException.class, () -> ExcelCustomXmlController.schemaXml(mapping, null));

      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMap ctMap =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMap.Factory.newInstance();
      ctMap.setID(9L);
      ctMap.setName("FakeMap");
      ctMap.setRootElement("Root");
      ctMap.setSchemaID("Schema9");
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSchema ctSchema =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSchema.Factory.newInstance();
      var fakeMap =
          ExcelCustomXmlControllerTestSupport.fakeMap(ctMap, ctSchema, null, List.of(), List.of());
      ExcelCustomXmlMappingSnapshot snapshot = ExcelCustomXmlController.snapshot(fakeMap);
      assertNull(snapshot.schemaNamespace());
      assertNull(snapshot.schemaLanguage());
      assertNull(snapshot.schemaReference());
      assertNull(snapshot.schemaXml());
    }
  }

  @SuppressWarnings("PMD.SignatureDeclareThrowsException")
  private static void applySeriesOptionBranches(XSSFChart chart) throws Exception {
    var categories = XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Feb"});
    var values = XDDFDataSourcesFactory.fromArray(new Double[] {10d, 12d});
    ExcelChartDefinition.Series markerSeries =
        new ExcelChartDefinition.Series(
            null,
            new ExcelChartDefinition.DataSource.StringLiteral(List.of("Jan", "Feb")),
            new ExcelChartDefinition.DataSource.NumericLiteral(List.of(10d, 12d)),
            true,
            ExcelChartMarkerStyle.SQUARE,
            (short) 9,
            20L);
    ExcelChartDefinition.Series noOptionsSeries =
        new ExcelChartDefinition.Series(
            null,
            new ExcelChartDefinition.DataSource.StringLiteral(List.of("Jan", "Feb")),
            new ExcelChartDefinition.DataSource.NumericLiteral(List.of(10d, 12d)),
            null,
            null,
            null,
            null);

    XDDFLineChartData lineData =
        (XDDFLineChartData)
            chart.createData(
                ChartTypes.LINE,
                chart.createCategoryAxis(AxisPosition.BOTTOM),
                chart.createValueAxis(AxisPosition.LEFT));
    XDDFLineChartData.Series lineSeries =
        (XDDFLineChartData.Series) lineData.addSeries(categories, values);
    ExcelChartPlotMutationSupport.applySeriesOptions(lineSeries, markerSeries);
    assertEquals("square", lineSeries.getCTLineSer().getMarker().getSymbol().getVal().toString());
    assertEquals(9L, lineSeries.getCTLineSer().getMarker().getSize().getVal());
    ExcelChartPlotMutationSupport.applySeriesOptions(lineSeries, noOptionsSeries);

    XDDFLine3DChartData line3DData =
        (XDDFLine3DChartData)
            chart.createData(
                ChartTypes.LINE3D,
                chart.createCategoryAxis(AxisPosition.TOP),
                chart.createValueAxis(AxisPosition.RIGHT));
    XDDFLine3DChartData.Series line3DSeries =
        (XDDFLine3DChartData.Series) line3DData.addSeries(categories, values);
    ExcelChartPlotMutationSupport.applySeriesOptions(line3DSeries, markerSeries);
    ExcelChartPlotMutationSupport.applySeriesOptions(line3DSeries, noOptionsSeries);

    XDDFScatterChartData scatterData =
        (XDDFScatterChartData)
            chart.createData(
                ChartTypes.SCATTER,
                chart.createValueAxis(AxisPosition.BOTTOM),
                chart.createValueAxis(AxisPosition.LEFT));
    XDDFScatterChartData.Series scatterSeries =
        (XDDFScatterChartData.Series) scatterData.addSeries(values, values);
    ExcelChartPlotMutationSupport.applySeriesOptions(scatterSeries, markerSeries);
    ExcelChartPlotMutationSupport.applySeriesOptions(scatterSeries, noOptionsSeries);

    ExcelChartDefinition.Series explodedSeries =
        new ExcelChartDefinition.Series(
            null,
            new ExcelChartDefinition.DataSource.StringLiteral(List.of("Jan", "Feb")),
            new ExcelChartDefinition.DataSource.NumericLiteral(List.of(10d, 12d)),
            null,
            null,
            null,
            25L);
    XDDFPieChartData pieData = (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);
    XDDFPieChartData.Series pieSeries =
        (XDDFPieChartData.Series) pieData.addSeries(categories, values);
    ExcelChartPlotMutationSupport.applySeriesOptions(pieSeries, explodedSeries);
    assertTrue(pieSeries.getCTPieSer().isSetExplosion());
    ExcelChartPlotMutationSupport.applySeriesOptions(pieSeries, noOptionsSeries);

    XDDFPie3DChartData pie3DData =
        (XDDFPie3DChartData) chart.createData(ChartTypes.PIE3D, null, null);
    XDDFPie3DChartData.Series pie3DSeries =
        (XDDFPie3DChartData.Series) pie3DData.addSeries(categories, values);
    ExcelChartPlotMutationSupport.applySeriesOptions(pie3DSeries, explodedSeries);
    assertTrue(pie3DSeries.getCTPieSer().isSetExplosion());
    ExcelChartPlotMutationSupport.applySeriesOptions(pie3DSeries, noOptionsSeries);

    var doughnutData =
        (org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData)
            chart.createData(ChartTypes.DOUGHNUT, null, null);
    var doughnutSeries =
        (org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData.Series)
            doughnutData.addSeries(categories, values);
    ExcelChartPlotMutationSupport.applySeriesOptions(doughnutSeries, explodedSeries);
    assertTrue(doughnutSeries.getCTPieSer().isSetExplosion());
    ExcelChartPlotMutationSupport.applySeriesOptions(doughnutSeries, noOptionsSeries);
  }

  @SuppressWarnings("PMD.SignatureDeclareThrowsException")
  private static void applySeriesTitleAndChartTitleBranches(XSSFChart chart, XSSFSheet sheet)
      throws Exception {
    var categories = XDDFDataSourcesFactory.fromArray(new String[] {"Jan", "Feb"});
    var values = XDDFDataSourcesFactory.fromArray(new Double[] {10d, 12d});
    CellReference titleCell = new CellReference(sheet.getSheetName(), 1, 1, true, true);

    XDDFBarChartData barData =
        (XDDFBarChartData)
            chart.createData(
                ChartTypes.BAR,
                chart.createCategoryAxis(AxisPosition.BOTTOM),
                chart.createValueAxis(AxisPosition.LEFT));
    XDDFBarChartData.Series barSeries =
        (XDDFBarChartData.Series) barData.addSeries(categories, values);
    barSeries.setTitle("Old");
    ExcelChartMutationSupport.applySeriesTitle(barSeries, new PreparedSeriesTitleText("Bar Title"));
    ExcelChartMutationSupport.applySeriesTitle(
        barSeries, new PreparedSeriesTitleFormula("Series", titleCell));
    ExcelChartMutationSupport.applySeriesTitle(barSeries, new PreparedSeriesTitleNone());
    assertFalse(barSeries.getCTBarSer().isSetTx());

    XDDFLineChartData lineData =
        (XDDFLineChartData)
            chart.createData(
                ChartTypes.LINE,
                chart.createCategoryAxis(AxisPosition.TOP),
                chart.createValueAxis(AxisPosition.RIGHT));
    XDDFLineChartData.Series lineSeries =
        (XDDFLineChartData.Series) lineData.addSeries(categories, values);
    lineSeries.setTitle("Old");
    ExcelChartMutationSupport.applySeriesTitle(
        lineSeries, new PreparedSeriesTitleText("Line Title"));
    ExcelChartMutationSupport.applySeriesTitle(
        lineSeries, new PreparedSeriesTitleFormula("Series", titleCell));
    ExcelChartMutationSupport.applySeriesTitle(lineSeries, new PreparedSeriesTitleNone());
    assertFalse(lineSeries.getCTLineSer().isSetTx());

    XDDFPieChartData pieData = (XDDFPieChartData) chart.createData(ChartTypes.PIE, null, null);
    XDDFPieChartData.Series pieSeries =
        (XDDFPieChartData.Series) pieData.addSeries(categories, values);
    pieSeries.setTitle("Old");
    ExcelChartMutationSupport.applySeriesTitle(pieSeries, new PreparedSeriesTitleText("Pie Title"));
    ExcelChartMutationSupport.applySeriesTitle(
        pieSeries, new PreparedSeriesTitleFormula("Series", titleCell));
    ExcelChartMutationSupport.applySeriesTitle(pieSeries, new PreparedSeriesTitleNone());
    assertFalse(pieSeries.getCTPieSer().isSetTx());

    chart.setTitleText("Old title");
    ExcelChartMutationSupport.applyChartTitleFormula(chart, "Budget", titleCell);
    ExcelChartMutationSupport.applyChartTitleFormula(chart, "Budget 2", titleCell);
    assertFalse(chart.getCTChart().getTitle().getTx().isSetRich());
    assertEquals(
        "Budget 2",
        chart.getCTChart().getTitle().getTx().getStrRef().getStrCache().getPtArray(0).getV());
  }

  private static ExcelSignatureLineDefinition signatureDefinition(
      String name, ExcelDrawingAnchor.TwoCell anchor) {
    return new ExcelSignatureLineDefinition(
        name,
        anchor,
        false,
        "Review before signing.",
        "Ada Lovelace",
        "Finance",
        "ada@example.com",
        null,
        "invalid",
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(new byte[] {1}));
  }

  private static ExcelChartDefinition.Axis axis(ExcelChartAxisKind kind) {
    ExcelChartAxisPosition position =
        switch (kind) {
          case CATEGORY -> ExcelChartAxisPosition.BOTTOM;
          case DATE -> ExcelChartAxisPosition.TOP;
          case SERIES -> ExcelChartAxisPosition.RIGHT;
          case VALUE -> ExcelChartAxisPosition.LEFT;
        };
    return new ExcelChartDefinition.Axis(kind, position, ExcelChartAxisCrosses.AUTO_ZERO, true);
  }

  private static void seedChartData(XSSFSheet sheet) {
    sheet.createRow(0).createCell(0).setCellValue("Month");
    sheet.getRow(0).createCell(1).setCellValue("Plan");
    sheet.createRow(1).createCell(0).setCellValue("Jan");
    sheet.getRow(1).createCell(1).setCellValue(10d);
    sheet.createRow(2).createCell(0).setCellValue("Feb");
    sheet.getRow(2).createCell(1).setCellValue(12d);
    sheet.createRow(3).createCell(0).setCellValue("Mar");
    sheet.getRow(3).createCell(1).setCellValue(15d);
  }

  private static void snapshotPlotDefaultsForUnsetGroupingAndShape(
      XSSFDrawing drawing,
      org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource categories,
      org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource<Double> values) {
    XSSFChart area3DChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 37, 1, 43, 10));
    XDDFArea3DChartData area3DData =
        (XDDFArea3DChartData)
            area3DChart.createData(
                ChartTypes.AREA3D,
                area3DChart.createCategoryAxis(AxisPosition.BOTTOM),
                area3DChart.createValueAxis(AxisPosition.LEFT));
    area3DData.addSeries(categories, values);
    area3DChart.plot(area3DData);
    assertInstanceOf(
        ExcelChartSnapshot.Area3D.class,
        ExcelChartSnapshotSupport.snapshotChart(area3DChart, area3DChart.getGraphicFrame())
            .plots()
            .getFirst());

    XSSFChart bar3DChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 44, 1, 50, 10));
    XDDFBar3DChartData bar3DData =
        (XDDFBar3DChartData)
            bar3DChart.createData(
                ChartTypes.BAR3D,
                bar3DChart.createCategoryAxis(AxisPosition.BOTTOM),
                bar3DChart.createValueAxis(AxisPosition.LEFT));
    bar3DData.addSeries(categories, values);
    bar3DChart.plot(bar3DData);
    assertInstanceOf(
        ExcelChartSnapshot.Bar3D.class,
        ExcelChartSnapshotSupport.snapshotChart(bar3DChart, bar3DChart.getGraphicFrame())
            .plots()
            .getFirst());

    XSSFChart line3DChart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 51, 1, 57, 10));
    XDDFLine3DChartData line3DData =
        (XDDFLine3DChartData)
            line3DChart.createData(
                ChartTypes.LINE3D,
                line3DChart.createCategoryAxis(AxisPosition.BOTTOM),
                line3DChart.createValueAxis(AxisPosition.LEFT));
    line3DData.addSeries(categories, values);
    line3DChart.plot(line3DData);
    assertInstanceOf(
        ExcelChartSnapshot.Line3D.class,
        ExcelChartSnapshotSupport.snapshotChart(line3DChart, line3DChart.getGraphicFrame())
            .plots()
            .getFirst());
  }

  private static List<CTShape> signatureShapes(XSSFVMLDrawing vmlDrawing) {
    List<CTShape> shapes = new java.util.ArrayList<>();
    try (XmlCursor cursor = vmlDrawing.getDocument().getXml().newCursor()) {
      for (boolean found = cursor.toFirstChild(); found; found = cursor.toNextSibling()) {
        XmlObject object = cursor.getObject();
        if (object instanceof CTShape shape && shape.sizeOfSignaturelineArray() > 0) {
          shapes.add(shape);
        }
      }
    }
    return List.copyOf(shapes);
  }

  private static ExcelChartMarkerStyle invokeMarkerStyle(CTMarker marker) {
    return ExcelChartPlotSnapshotSupport.markerStyle(marker).orElse(null);
  }

  private static Short invokeMarkerSize(CTMarker marker) {
    return ExcelChartPlotSnapshotSupport.markerSize(marker);
  }

  private static void restoreProperty(String name, String value) {
    if (value == null) {
      System.clearProperty(name);
    } else {
      System.setProperty(name, value);
    }
  }

  /** Minimal unsupported chart-data stub for residual unsupported-plot coverage. */
  private static final class UnsupportedChartData extends XDDFChartData {
    private UnsupportedChartData() {
      super(null);
    }

    @Override
    protected void removeCTSeries(int index) {}

    @Override
    public void setVaryColors(Boolean varyColors) {}

    @Override
    public Series addSeries(
        org.apache.poi.xddf.usermodel.chart.XDDFDataSource<?> category,
        org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource<? extends Number> values) {
      return null;
    }
  }
}
