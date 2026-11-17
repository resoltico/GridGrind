package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import java.util.List;
import java.util.Optional;

/** Dense authored-chart fixtures assert declarations before comparing their required point text. */
final class ChartCacheAssertions {
  private ChartCacheAssertions() {}

  static String singleStringValue(ChartReport.StringCache cache) {
    List<String> values = denseStringValues(cache);
    assertEquals(1, values.size());
    return values.getFirst();
  }

  static List<String> denseStringValues(ChartReport.StringCache cache) {
    var stored = assertInstanceOf(ChartReport.StringCache.Stored.class, cache);
    assertEquals(Optional.of((long) stored.points().size()), stored.declaredCount());
    for (int index = 0; index < stored.points().size(); index++) {
      assertEquals(Optional.of((long) index), stored.points().get(index).index());
    }
    return stored.points().stream().map(point -> point.valueText().orElseThrow()).toList();
  }

  static List<String> denseNumericValues(ChartReport.NumericCache cache) {
    var stored = assertInstanceOf(ChartReport.NumericCache.Stored.class, cache);
    assertEquals(Optional.of((long) stored.points().size()), stored.declaredCount());
    for (int index = 0; index < stored.points().size(); index++) {
      assertEquals(Optional.of((long) index), stored.points().get(index).index());
    }
    return stored.points().stream().map(point -> point.valueText().orElseThrow()).toList();
  }
}
