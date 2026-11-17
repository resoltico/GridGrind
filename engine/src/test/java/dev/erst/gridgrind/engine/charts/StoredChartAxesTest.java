package dev.erst.gridgrind.engine.charts;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlException;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt;

/** Native axis crossing and visibility observation never invokes creating XDDF getters. */
class StoredChartAxesTest {
  @ParameterizedTest
  @ValueSource(strings = {"2.5", "2.500e0", "INF", "-INF", "NaN"})
  void completeChartReportsExactNativeNumericCrossingWithoutMutation(String lexical)
      throws IOException {
    try (var workbook = new XSSFWorkbook()) {
      var chart =
          workbook
              .createSheet("Native")
              .createDrawingPatriarch()
              .createChart(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 4, 4));
      chart.createData(
          ChartTypes.BAR,
          chart.createCategoryAxis(AxisPosition.BOTTOM),
          chart.createValueAxis(AxisPosition.LEFT));
      var axis = chart.getCTChart().getPlotArea().getValAxArray(0);
      axis.unsetCrosses();
      axis.unsetDelete();
      axis.addNewCrossesAt();
      try (var cursor = axis.getCrossesAt().newCursor()) {
        cursor.setAttributeText(new javax.xml.namespace.QName("val"), lexical);
      }
      assertTrue(chart.getCTChartSpace().validate());
      String before = chart.getCTChartSpace().xmlText();
      var expectedCrossing = new ChartReport.AxisCrossing.Value(lexical);
      for (int repeat = 0; repeat < 2; repeat++) {
        var plot =
            assertInstanceOf(
                ChartReport.Bar.class,
                ChartAdapter.snapshotChart(chart, chart.getGraphicFrame()).plots().getFirst());
        assertEquals(expectedCrossing, plot.axes().get(1).crosses());
        assertTrue(plot.axes().get(1).visible());
        ChartAdapter.snapshotChartDrawingObject(chart, chart.getGraphicFrame());
      }
      assertFalse(axis.isSetCrosses());
      assertFalse(axis.isSetDelete());
      assertEquals(before, chart.getCTChartSpace().xmlText());
    }
  }

  @ParameterizedTest
  @ValueSource(strings = {"catAx", "dateAx", "serAx", "valAx"})
  void allNativeAxisFamiliesKeepMissingModeValueAndDefaultVisibility(String family)
      throws XmlException {
    var area = area(family, "");
    assertTrue(
        area
            .selectChildren(
                new javax.xml.namespace.QName(
                    "http://schemas.openxmlformats.org/drawingml/2006/chart", family))[0]
            .validate());
    var reference = CTUnsignedInt.Factory.newInstance();
    reference.setVal(1);
    var reader = new StoredChartReader(Optional.of("/xl/charts/chart1.xml"));
    String before = area.xmlText();
    var missing = new StoredChartAxes(area, reader).selected(List.of(reference)).getFirst();
    assertEquals(new ChartReport.AxisCrossing.Missing(), missing.crosses());
    assertTrue(missing.visible());
    assertEquals(before, area.xmlText());
    var mode =
        new StoredChartAxes(area(family, "<c:crosses val='min'/>"), reader)
            .selected(List.of(reference))
            .getFirst();
    assertEquals(new ChartReport.AxisCrossing.Mode(ExcelChartAxisCrosses.MIN), mode.crosses());
    var value =
        new StoredChartAxes(area(family, "<c:crossesAt val='2.5'/>"), reader)
            .selected(List.of(reference))
            .getFirst();
    assertEquals(new ChartReport.AxisCrossing.Value("2.5"), value.crosses());
    reference.setVal(99);
    assertEquals(List.of(), new StoredChartAxes(area, reader).selected(List.of(reference)));
  }

  @ParameterizedTest
  @ValueSource(
      strings = {
        "<c:crosses val='min'/><c:crossesAt val='2.5'/>",
        "<c:crossesAt val='2.5'/><c:crossesAt val='3'/>",
        "<c:crosses val='min'/><c:crosses val='max'/>",
        "<c:crossesAt val='not-a-number'/>",
        "<c:crosses val='unexpected'/>"
      })
  void mixedRepeatedAndInvalidNativeCrossingsRefuseWithoutMutation(String crossing)
      throws XmlException {
    var area = area("valAx", crossing);
    var reference = CTUnsignedInt.Factory.newInstance();
    reference.setVal(1);
    var reader = new StoredChartReader(Optional.of("/xl/charts/chart1.xml"));
    String before = area.xmlText();
    var refusal =
        assertThrows(
            WorkbookContentRefusalException.class,
            () ->
                reader.read(
                    "chart.axis",
                    () -> new StoredChartAxes(area, reader).selected(List.of(reference))));
    assertEquals(WorkbookContentIssue.INVALID_NATIVE_VALUE, refusal.issue());
    assertEquals(Optional.of("/xl/charts/chart1.xml"), refusal.location().orElseThrow().partName());
    assertEquals(before, area.xmlText());
  }

  @Test
  void missingRequiredPositionAndRepeatedReportedSingletonAreRefused() throws XmlException {
    var reference = CTUnsignedInt.Factory.newInstance();
    reference.setVal(1);
    var reader = new StoredChartReader(Optional.empty());
    var missing = area("valAx", "");
    try (var cursor = missing.getValAxArray(0).getAxPos().newCursor()) {
      cursor.removeXml();
    }
    assertThrows(
        WorkbookContentRefusalException.class,
        () -> new StoredChartAxes(missing, reader).selected(List.of(reference)));
    var repeated = area("valAx", "");
    repeated.getValAxArray(0).addNewDelete();
    repeated.getValAxArray(0).addNewDelete();
    assertThrows(
        WorkbookContentRefusalException.class,
        () -> new StoredChartAxes(repeated, reader).selected(List.of(reference)));
  }

  private static CTPlotArea area(String family, String crossing) throws XmlException {
    var area =
        CTPlotArea.Factory.parse(
            "<xml-fragment xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart'><c:"
                + family
                + "><c:axId val='1'/><c:scaling/><c:axPos val='b'/><c:crossAx val='2'/>"
                + crossing
                + "</c:"
                + family
                + "></xml-fragment>");
    return area;
  }
}
