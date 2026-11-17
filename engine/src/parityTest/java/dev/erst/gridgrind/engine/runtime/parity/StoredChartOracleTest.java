package dev.erst.gridgrind.engine.runtime.parity;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;

/** Literal expectations distinguish raw physical XML declarations from dense/live XDDF values. */
class StoredChartOracleTest {
  @Test
  void numericOracleKeepsOrderGapsBlankErrorAndMissingDeclarations() {
    var source = CTBarSer.Factory.newInstance().addNewVal();
    var reference = source.addNewNumRef();
    reference.setF("Unresolvable!$A$1:$A$5");
    var cache = reference.addNewNumCache();
    cache.addNewPtCount().setVal(9);
    cache.setFormatCode("");
    var error = cache.addNewPt();
    error.setIdx(4);
    error.setV("#N/A");
    error.setFormatCode("");
    var blank = cache.addNewPt();
    blank.setIdx(0);
    blank.setV("");
    cache.addNewPt();
    String before = source.xmlText();
    var actual =
        assertInstanceOf(
            ChartReport.DataSource.NumericReference.class, StoredChartOracle.source(source));
    assertEquals(
        new ChartReport.NumericCache.Stored(
            Optional.of(9L),
            Optional.of(""),
            List.of(
                new ChartReport.NumericPoint(Optional.of(4L), Optional.of("#N/A"), Optional.of("")),
                new ChartReport.NumericPoint(Optional.of(0L), Optional.of(""), Optional.empty()),
                new ChartReport.NumericPoint(
                    Optional.empty(), Optional.empty(), Optional.empty()))),
        actual.cache());
    assertEquals(before, source.xmlText());
    reference.unsetNumCache();
    assertEquals(
        new ChartReport.DataSource.NumericReference(
            Optional.of("Unresolvable!$A$1:$A$5"), new ChartReport.NumericCache.Missing()),
        StoredChartOracle.source(source));
  }

  @Test
  void stringAndTitleOraclePreserveRawCountAndOptionalFormula() {
    var series = CTBarSer.Factory.newInstance();
    var literal = series.addNewCat().addNewStrLit();
    literal.addNewPtCount().setVal(7);
    var point = literal.addNewPt();
    point.setIdx(3);
    point.setV("");
    assertEquals(
        new ChartReport.DataSource.StringLiteral(
            new ChartReport.StringCache.Stored(
                Optional.of(7L),
                List.of(new ChartReport.StringPoint(Optional.of(3L), Optional.of(""))))),
        StoredChartOracle.source(series.getCat()));
    var title = series.addNewTx().addNewStrRef();
    assertEquals(
        new ChartReport.Title.Formula(Optional.empty(), new ChartReport.StringCache.Missing()),
        StoredChartOracle.formula(title));
    title.setF("");
    title.addNewStrCache().addNewPtCount().setVal(0);
    assertEquals(
        new ChartReport.Title.Formula(
            Optional.of(""), new ChartReport.StringCache.Stored(Optional.of(0L), List.of())),
        StoredChartOracle.formula(title));
  }

  @ParameterizedTest
  @ValueSource(strings = {"2.5", "2.500e0", "INF", "-INF", "NaN"})
  void axisOracleReadsRawCrossingAndAbsentVisibilityWithoutWrapperDefaults(String value) {
    var area = org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea.Factory.newInstance();
    var axis = area.addNewValAx();
    axis.addNewAxId().setVal(1);
    axis.addNewAxPos().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STAxPos.L);
    axis.addNewCrossesAt();
    try (var cursor = axis.getCrossesAt().newCursor()) {
      cursor.setAttributeText(new javax.xml.namespace.QName("val"), value);
    }
    var reference =
        org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt.Factory.newInstance();
    reference.setVal(1);
    String before = area.xmlText();
    var observed = StoredChartOracle.axes(area, List.of(reference)).getFirst();
    assertEquals(new ChartReport.AxisCrossing.Value(value), observed.crosses());
    assertTrue(observed.visible());
    assertEquals(before, area.xmlText());
    axis.unsetCrossesAt();
    assertEquals(
        new ChartReport.AxisCrossing.Missing(),
        StoredChartOracle.axes(area, List.of(reference)).getFirst().crosses());
  }
}
