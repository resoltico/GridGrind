package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

/** Verifies log parsing for run summaries emitted by the local Jazzer operator layer. */
class JazzerReportSupportTest {
  /** Parses libFuzzer summaries that report corpus bytes using kilobyte units. */
  @Test
  void parseActiveFuzzMetrics_parsesKilobyteCorpusSummary() {
    String log =
        """
        #9299\tDONE   cov: 617 ft: 1704 corp: 295/17Kb lim: 363 exec/s: 845 rss: 983Mb
        Done 9299 runs in 11 second(s)
        """;

    RunMetrics.ActiveFuzzMetrics metrics = JazzerReportSupport.parseActiveFuzzMetrics(log);

    assertEquals(9299L, metrics.executions());
    assertEquals(617, metrics.coverage());
    assertEquals(1704, metrics.features());
    assertEquals(295, metrics.corpusEntries());
    assertEquals(17L * 1024L, metrics.corpusBytes());
    assertEquals(363, metrics.maxInputBytes());
    assertEquals(845, metrics.executionsPerSecond());
    assertEquals(983, metrics.rssMegabytes());
  }
}
