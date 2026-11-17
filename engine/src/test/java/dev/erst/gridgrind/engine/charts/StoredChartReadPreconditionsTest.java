package dev.erst.gridgrind.engine.charts;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import java.io.IOException;
import java.util.List;
import javax.xml.namespace.QName;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

/** Refuse only omissions that the pinned XDDF read code would repair by writing native XML. */
class StoredChartReadPreconditionsTest {
  private static final String CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";

  @ParameterizedTest
  @EnumSource(
      value = ChartTypes.class,
      names = {
        "AREA",
        "AREA3D",
        "BAR",
        "BAR3D",
        "LINE",
        "LINE3D",
        "RADAR",
        "SCATTER",
        "SURFACE",
        "SURFACE3D"
      })
  void zeroStoredAxisReferencesRefuseBeforeAnyReadEntryCanInventThem(ChartTypes type)
      throws IOException {
    try (var workbook = new XSSFWorkbook()) {
      var chart = chart(workbook);
      var data =
          chart.createData(
              type,
              chart.createCategoryAxis(AxisPosition.BOTTOM),
              chart.createValueAxis(AxisPosition.LEFT));
      remove(plot(chart, type), "axId");
      assertAllReadEntriesRefuseUnchanged(chart, data);
    }
  }

  @ParameterizedTest
  @EnumSource(
      value = ChartTypes.class,
      names = {"BAR", "BAR3D", "SCATTER"})
  void missingRequiredDirectionOrScatterStyleRefusesBeforeDefaultCreation(ChartTypes type)
      throws IOException {
    try (var workbook = new XSSFWorkbook()) {
      var chart = chart(workbook);
      var data =
          chart.createData(
              type,
              chart.createCategoryAxis(AxisPosition.BOTTOM),
              chart.createValueAxis(AxisPosition.LEFT));
      remove(plot(chart, type), type == ChartTypes.SCATTER ? "scatterStyle" : "barDir");
      assertAllReadEntriesRefuseUnchanged(chart, data);
    }
  }

  @ParameterizedTest
  @EnumSource(
      value = ChartTypes.class,
      names = {"AREA", "AREA3D", "BAR", "BAR3D"})
  void schemaValidOptionalGroupingRemainsAbsentAndZeroSeriesStayObservable(ChartTypes type)
      throws IOException {
    try (var workbook = new XSSFWorkbook()) {
      var chart = chart(workbook);
      var data =
          chart.createData(
              type,
              chart.createCategoryAxis(AxisPosition.BOTTOM),
              chart.createValueAxis(AxisPosition.LEFT));
      remove(plot(chart, type), "grouping");
      assertTrue(chart.getCTChartSpace().validate());
      String before = chart.getCTChartSpace().xmlText();
      assertEquals(1, ChartAdapter.snapshotChart(chart, chart.getGraphicFrame()).plots().size());
      assertEquals(
          List.of(ChartAdapter.plotTypeToken(data)),
          ChartAdapter.snapshotChartDrawingObject(chart, chart.getGraphicFrame()).plotTypeTokens());
      assertEquals(1, ChartAdapter.snapshotPlots(chart, List.of(data)).size());
      assertEquals(before, chart.getCTChartSpace().xmlText());
    }
  }

  @Test
  void invalidAndRepeatedRequiredNativeDirectionsRefuseWithoutSingletonLoss() throws IOException {
    try (var workbook = new XSSFWorkbook()) {
      var chart = chart(workbook);
      var data =
          chart.createData(
              ChartTypes.BAR,
              chart.createCategoryAxis(AxisPosition.BOTTOM),
              chart.createValueAxis(AxisPosition.LEFT));
      var bar = chart.getCTChart().getPlotArea().getBarChartArray(0);
      try (var cursor = bar.getBarDir().newCursor()) {
        cursor.setAttributeText(new QName("val"), "not-a-direction");
      }
      assertAllReadEntriesRefuseUnchanged(chart, data);
      bar.getBarDir().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir.BAR);
      bar.addNewBarDir().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir.COL);
      assertAllReadEntriesRefuseUnchanged(chart, data);
    }
  }

  @Test
  void laterMalformedPlotIsCheckedBeforeAnyConstructorRuns() throws IOException {
    try (var workbook = new XSSFWorkbook()) {
      var chart = chart(workbook);
      var category = chart.createCategoryAxis(AxisPosition.BOTTOM);
      var value = chart.createValueAxis(AxisPosition.LEFT);
      chart.createData(ChartTypes.AREA, category, value);
      var bar = chart.createData(ChartTypes.BAR, category, value);
      remove(plot(chart, ChartTypes.BAR), "barDir");
      assertAllReadEntriesRefuseUnchanged(chart, bar);
    }
  }

  private static void assertAllReadEntriesRefuseUnchanged(XSSFChart chart, XDDFChartData data) {
    String before = chart.getCTChartSpace().xmlText();
    List<Runnable> readers =
        List.of(
            () -> ChartAdapter.snapshotChart(chart, chart.getGraphicFrame()),
            () ->
                ChartAdapter.snapshotChartDrawingObject(chart, chart.getGraphicFrame())
                    .plotTypeTokens(),
            () -> ChartAdapter.snapshotPlots(chart, List.of(data)));
    for (Runnable read : readers) {
      assertEquals(
          WorkbookContentIssue.INVALID_NATIVE_VALUE,
          assertThrows(WorkbookContentRefusalException.class, read::run).issue());
      assertEquals(before, chart.getCTChartSpace().xmlText());
    }
  }

  private static XSSFChart chart(XSSFWorkbook workbook) {
    return workbook
        .createSheet("Native")
        .createDrawingPatriarch()
        .createChart(new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 4, 4));
  }

  private static XmlObject plot(XSSFChart chart, ChartTypes type) {
    String name =
        switch (type) {
          case AREA -> "areaChart";
          case AREA3D -> "area3DChart";
          case BAR -> "barChart";
          case BAR3D -> "bar3DChart";
          case LINE -> "lineChart";
          case LINE3D -> "line3DChart";
          case RADAR -> "radarChart";
          case SCATTER -> "scatterChart";
          case SURFACE -> "surfaceChart";
          case SURFACE3D -> "surface3DChart";
          default -> throw new IllegalArgumentException("Fixture requires an axis plot");
        };
    return chart.getCTChart().getPlotArea().selectChildren(new QName(CHART, name))[0];
  }

  private static void remove(XmlObject parent, String name) {
    for (XmlObject child : parent.selectChildren(new QName(CHART, name))) {
      try (var cursor = child.newCursor()) {
        cursor.removeXml();
      }
    }
  }
}
