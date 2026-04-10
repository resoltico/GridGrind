package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

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

  /** Resolves project-relative promoted-input paths back to absolute files inside the workspace. */
  @Test
  void promotedInputPaths_resolveProjectRelativeMetadataEntries(@TempDir Path projectDirectory)
      throws IOException {
    Path promotedInputPath =
        JazzerHarness.protocolRequest().inputDirectory(projectDirectory).resolve("seed.json");
    Files.createDirectories(promotedInputPath.getParent());
    Files.writeString(promotedInputPath, "{}");

    Path metadataPath =
        JazzerHarness.protocolRequest().promotedMetadataDirectory(projectDirectory)
            .resolve("seed.json");
    Files.createDirectories(metadataPath.getParent());
    Path replayTextPath = metadataPath.resolveSibling("seed.txt");
    Files.writeString(replayTextPath, "Replay Result" + System.lineSeparator());
    JazzerJson.write(
        metadataPath,
        new PromotionMetadata(
            JazzerHarness.protocolRequest().key(),
            PromotionMetadata.relativizePath(projectDirectory, promotedInputPath),
            PromotionMetadata.relativizePath(projectDirectory, promotedInputPath),
            "SUCCESS",
            new ReplayExpectation(
                "SUCCESS",
                new ProtocolRequestDetails(2, "PARSED", "NEW", "NONE", 0, Map.of(), Map.of(), 0, Map.of())),
            "2026-04-06T00:00:00Z",
            PromotionMetadata.relativizePath(projectDirectory, replayTextPath)));

    assertEquals(
        Set.of(promotedInputPath.toAbsolutePath().normalize()),
        JazzerReportSupport.promotedInputPaths(projectDirectory));
    assertEquals(
        List.of(),
        JazzerReportSupport.orphanedInputs(projectDirectory, JazzerHarness.protocolRequest()));
  }
}
