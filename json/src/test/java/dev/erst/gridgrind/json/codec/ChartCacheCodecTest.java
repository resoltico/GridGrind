package dev.erst.gridgrind.json.codec;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingCoordinateReport;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** The exchange format preserves stored-cache presence and raw point declarations. */
class ChartCacheCodecTest {
  @Test
  void rawBlankErrorMissingAndCountDeclarationsRoundTripWithoutLiveValues() {
    ChartReport.DataSource expected =
        new ChartReport.DataSource.NumericReference(
            Optional.of("Missing!$B$2"),
            new ChartReport.NumericCache.Stored(
                Optional.of(9L),
                Optional.of(""),
                List.of(
                    new ChartReport.NumericPoint(
                        Optional.of(4L), Optional.of("#N/A"), Optional.empty()),
                    new ChartReport.NumericPoint(Optional.of(0L), Optional.of(""), Optional.of("")),
                    new ChartReport.NumericPoint(
                        Optional.empty(), Optional.empty(), Optional.empty()))));
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    var json = mapper.valueToTree(expected);
    assertEquals("NUMERIC_REFERENCE", json.path("type").asString());
    assertFalse(json.has("cachedValues"));
    assertEquals("STORED", json.path("cache").path("type").asString());
    assertEquals(9L, json.path("cache").path("declaredCount").asLong());
    assertEquals("", json.path("cache").path("formatCode").asString());
    assertTrue(json.path("cache").has("formatCode"));
    var points = json.path("cache").path("points");
    assertEquals("#N/A", points.get(0).path("valueText").asString());
    assertEquals("", points.get(1).path("valueText").asString());
    assertTrue(points.get(1).has("valueText"));
    assertFalse(points.get(2).has("valueText"));
    assertFalse(points.get(2).has("index"));
    assertEquals(expected, mapper.treeToValue(json, ChartReport.DataSource.class));
  }

  @Test
  void literalSourceConcreteCacheRecordsRoundTripThroughTheSourceUnion() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    List<ChartReport.DataSource> values =
        List.of(
            new ChartReport.DataSource.StringLiteral(
                new ChartReport.StringCache.Stored(Optional.of(0L), List.of())),
            new ChartReport.DataSource.NumericLiteral(
                new ChartReport.NumericCache.Stored(
                    Optional.empty(),
                    Optional.of(""),
                    List.of(
                        new ChartReport.NumericPoint(
                            Optional.of(3L), Optional.of("#DIV/0!"), Optional.empty())))));
    for (ChartReport.DataSource value : values) {
      assertEquals(
          value, mapper.treeToValue(mapper.valueToTree(value), ChartReport.DataSource.class));
    }
  }

  @Test
  void formulaTitleDoesNotFlattenMissingAndEmptyStoredCaches() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    ChartReport.Title absent =
        new ChartReport.Title.Formula(Optional.empty(), new ChartReport.StringCache.Missing());
    ChartReport.Title empty =
        new ChartReport.Title.Formula(
            Optional.of(""), new ChartReport.StringCache.Stored(Optional.of(0L), List.of()));
    var absentJson = mapper.valueToTree(absent);
    var emptyJson = mapper.valueToTree(empty);
    assertFalse(absentJson.has("formula"));
    assertEquals("MISSING", absentJson.path("cache").path("type").asString());
    assertTrue(emptyJson.has("formula"));
    assertEquals("", emptyJson.path("formula").asString());
    assertEquals("STORED", emptyJson.path("cache").path("type").asString());
    assertEquals(absent, mapper.treeToValue(absentJson, ChartReport.Title.class));
    assertEquals(empty, mapper.treeToValue(emptyJson, ChartReport.Title.class));
  }

  @Test
  void observedCoordinatesRoundTripTheirDistinctUnitsAndExactText() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    List<DrawingCoordinateReport> coordinates =
        List.of(
            new DrawingCoordinateReport.Emu(-1),
            new DrawingCoordinateReport.Measured("-01.00cm"),
            new DrawingCoordinateReport.Pixels(17));
    for (DrawingCoordinateReport coordinate : coordinates) {
      assertEquals(
          coordinate,
          mapper.treeToValue(mapper.valueToTree(coordinate), DrawingCoordinateReport.class));
    }
    DrawingAnchorReport expected =
        new DrawingAnchorReport.Absolute(
            coordinates.getFirst(), coordinates.get(1), 0, 27273042316900L, null);
    assertEquals(
        expected, mapper.treeToValue(mapper.valueToTree(expected), DrawingAnchorReport.class));
  }

  @Test
  void storedAxisCrossingChoicesRoundTripWithoutNonfiniteJsonNumbers() {
    var mapper = GridGrindJsonMapperSupport.JSON_MAPPER;
    List<ChartReport.AxisCrossing> crossings =
        List.of(
            new ChartReport.AxisCrossing.Missing(),
            new ChartReport.AxisCrossing.Mode(
                dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses.MIN),
            new ChartReport.AxisCrossing.Value("NaN"),
            new ChartReport.AxisCrossing.Value("2.500e0"));
    for (ChartReport.AxisCrossing crossing : crossings) {
      assertEquals(
          crossing,
          mapper.treeToValue(mapper.valueToTree(crossing), ChartReport.AxisCrossing.class));
    }
  }
}
