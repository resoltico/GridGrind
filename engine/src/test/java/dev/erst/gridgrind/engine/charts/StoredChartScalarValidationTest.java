package dev.erst.gridgrind.engine.charts;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import java.io.IOException;
import java.util.Optional;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;

/** Known native numeric facets refuse before POI conversion or narrowing. */
class StoredChartScalarValidationTest {
  @ParameterizedTest
  @ValueSource(strings = {"87.5%", "501%"})
  void invalidGapIsAStableLocatedRefusalRatherThanNumberConversionFailure(String lexical)
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
      var width = chart.getCTChart().getPlotArea().getBarChartArray(0).addNewGapWidth();
      width.setVal("87%");
      assertTrue(width.xgetVal().validate());
      var good =
          assertInstanceOf(
              ChartReport.Bar.class,
              ChartAdapter.snapshotChart(chart, chart.getGraphicFrame()).plots().getFirst());
      assertEquals(Optional.of(87), good.gapWidth());
      width.setVal(lexical);
      String before = chart.getCTChartSpace().xmlText();
      var refusal =
          assertThrows(
              WorkbookContentRefusalException.class,
              () -> ChartAdapter.snapshotChart(chart, chart.getGraphicFrame()));
      assertEquals(WorkbookContentIssue.INVALID_NATIVE_VALUE, refusal.issue());
      assertEquals(Optional.of("chart.gapWidth"), refusal.location().orElseThrow().fieldName());
      assertEquals(before, chart.getCTChartSpace().xmlText());
    }
  }

  @ParameterizedTest
  @ValueSource(strings = {"gapDepth", "overlap", "firstSliceAng", "explosion", "size"})
  void actualNativeNumericLeafTypesAreValidated(String field) {
    XmlObject source;
    if ("size".equals(field)) {
      var series = CTLineSer.Factory.newInstance();
      series.addNewMarker().addNewSize().setVal((short) 1);
      source = series.getMarker();
    } else if ("explosion".equals(field)) {
      var series = CTPieSer.Factory.newInstance();
      series.addNewExplosion().setVal(-1);
      source = series;
    } else {
      var bar = CTBarChart.Factory.newInstance();
      if ("gapDepth".equals(field)) {
        var area =
            org.openxmlformats.schemas.drawingml.x2006.chart.CTArea3DChart.Factory.newInstance();
        area.addNewGapDepth().setVal("501%");
        source = area;
      } else if ("firstSliceAng".equals(field)) {
        var pie = org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart.Factory.newInstance();
        pie.addNewFirstSliceAng().setVal(361);
        source = pie;
      } else {
        bar.addNewOverlap().setVal("101%");
        source = bar;
      }
    }
    var reader = new StoredChartReader(Optional.empty());
    assertEquals(
        WorkbookContentIssue.INVALID_NATIVE_VALUE,
        assertThrows(
                WorkbookContentRefusalException.class, () -> reader.requireFields(source, field))
            .issue());
  }
}
