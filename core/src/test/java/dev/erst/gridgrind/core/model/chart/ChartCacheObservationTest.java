package dev.erst.gridgrind.core.model.chart;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.problem.ValueInvariantException;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Stored chart facts retain cache declarations rather than authored-data guarantees. */
class ChartCacheObservationTest {
  @Test
  void missingCacheAndStoredEmptyCacheRemainDifferentFacts() {
    assertNotEquals(
        new ChartReport.StringCache.Missing(),
        new ChartReport.StringCache.Stored(Optional.of(0L), List.of()));
    assertNotEquals(
        new ChartReport.NumericCache.Missing(),
        new ChartReport.NumericCache.Stored(Optional.empty(), Optional.empty(), List.of()));
    assertNotEquals(
        new ChartReport.DataSource.Missing(),
        new ChartReport.DataSource.StringReference(
            Optional.empty(), new ChartReport.StringCache.Missing()));
  }

  @Test
  void physicalPointOrderIndicesCountsAndRawNumericValuesAreIndependent() {
    var points =
        new ArrayList<>(
            List.of(
                new ChartReport.NumericPoint(
                    Optional.of(4L), Optional.of("#N/A"), Optional.empty()),
                new ChartReport.NumericPoint(
                    Optional.of(0L), Optional.of("12.7500"), Optional.of("")),
                new ChartReport.NumericPoint(Optional.of(2L), Optional.of(""), Optional.empty()),
                new ChartReport.NumericPoint(
                    Optional.of(2L), Optional.of("NaN"), Optional.of("General")),
                new ChartReport.NumericPoint(
                    Optional.empty(), Optional.empty(), Optional.empty())));
    var expected = List.copyOf(points);
    var cache = new ChartReport.NumericCache.Stored(Optional.of(9L), Optional.of(""), points);
    points.clear();
    assertEquals(Optional.of(9L), cache.declaredCount());
    assertEquals(Optional.of(""), cache.formatCode());
    assertEquals(expected, cache.points());
    assertThrows(UnsupportedOperationException.class, () -> cache.points().clear());
  }

  @Test
  void sourceAndTitleFormulasRetainAbsentAndBlankDeclarations() {
    var cache =
        new ChartReport.StringCache.Stored(
            Optional.of(2L),
            List.of(
                new ChartReport.StringPoint(Optional.of(1L), Optional.empty()),
                new ChartReport.StringPoint(Optional.of(0L), Optional.of(""))));
    var absent = new ChartReport.Title.Formula(Optional.empty(), cache);
    var blank = new ChartReport.Title.Formula(Optional.of(""), cache);
    assertNotEquals(absent, blank);
    assertEquals(Optional.empty(), absent.formula());
    assertEquals(Optional.of(""), blank.formula());
    assertEquals(
        cache,
        new ChartReport.DataSource.StringReference(Optional.of("Missing!A1"), cache).cache());
  }

  @Test
  void observedPlotsRetainEmptySeriesAndUnresolvedAxisDefinitions() {
    var plot =
        new ChartReport.Bar(
            false,
            ExcelChartBarDirection.COLUMN,
            ExcelChartBarGrouping.CLUSTERED,
            Optional.of(87),
            Optional.empty(),
            List.of(),
            List.of());
    assertEquals(List.of(), plot.axes());
    assertEquals(List.of(), plot.series());
    assertEquals(Optional.of(87), plot.gapWidth());
  }

  @ParameterizedTest
  @ValueSource(longs = {-1L, 4294967296L})
  void canonicalStoredDeclarationsRejectValuesOutsideTheNativeUnsignedRange(long value) {
    assertThrows(
        ValueInvariantException.class,
        () -> new ChartReport.StringPoint(Optional.of(value), Optional.empty()));
    assertThrows(
        ValueInvariantException.class,
        () -> new ChartReport.NumericCache.Stored(Optional.of(value), Optional.empty(), List.of()));
  }

  @Test
  void upperNativeDeclarationDoesNotRequireAllocatingThatManyPoints() {
    var point = new ChartReport.StringPoint(Optional.of(4294967295L), Optional.of("stored"));
    var cache = new ChartReport.StringCache.Stored(Optional.of(4294967295L), List.of(point));
    assertEquals(1, cache.points().size());
    assertEquals(Optional.of(4294967295L), cache.points().getFirst().index());
  }
}
