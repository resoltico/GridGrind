package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.jazzer.support.HarnessTelemetrySnapshot;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

/** Verifies the human-readable rendering used by the local Jazzer operator commands. */
class JazzerTextRendererTest {
  /** Distinguishes unexpected failures from expected-invalid and replay-clean raw artifacts. */
  @Test
  void renderSummary_distinguishesArtifactOutcomes() {
    LocalRunSummary summary =
        new LocalRunSummary(
            "xlsx-roundtrip",
            "XLSX Round Trip",
            "fuzzXlsxRoundTrip",
            RunMode.ACTIVE_FUZZING,
            "2026-03-26T12:00:00Z",
            "2026-03-26T12:00:10Z",
            10L,
            0,
            "SUCCESS",
            "/tmp/run.log",
            "/tmp/history",
            new CorpusStats(1L, 10L),
            new CorpusStats(2L, 20L),
            new RunMetrics.ActiveFuzzMetrics(100L, 10, 20, 2, 20L, 17, 50, 900),
            List.of(
                new HarnessTelemetrySnapshot(
                    "xlsx-roundtrip",
                    "XLSX Round Trip",
                    "2026-03-26T12:00:00Z",
                    "2026-03-26T12:00:10Z",
                    10L,
                    20L,
                    17L,
                    0L,
                    0L,
                    10L,
                    0L,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null)),
            List.of(
                new FindingArtifact(
                    "crash-clean",
                    "/tmp/crash-clean",
                    "SUCCESS",
                    "/tmp/crash-clean.json",
                    "/tmp/crash-clean.txt"),
                new FindingArtifact(
                    "crash-invalid",
                    "/tmp/crash-invalid",
                    "EXPECTED_INVALID",
                    "/tmp/crash-invalid.json",
                    "/tmp/crash-invalid.txt"),
                new FindingArtifact(
                    "crash-live",
                    "/tmp/crash-live",
                    "UNEXPECTED_FAILURE",
                    "/tmp/crash-live.json",
                    "/tmp/crash-live.txt")));

    String summaryText = JazzerTextRenderer.renderSummary(summary);
    String statusText = JazzerTextRenderer.renderStatus(List.of(summary));

    assertTrue(
        summaryText.contains(
            "Findings: 1 active, 1 expected-invalid artifacts, 1 replay-clean artifacts"));
    assertTrue(statusText.contains("findings=1, expected-invalid=1, replay-clean=1"));
  }

  /** Preserves relative input paths so committed replay text does not hard-code one workspace. */
  @Test
  void renderReplay_preservesRelativeInputPaths() {
    ReplayOutcome outcome =
        new ReplayOutcome.Success(
            "protocol-request",
            new ProtocolRequestDetails(
                2,
                "PARSED",
                "NEW",
                "NONE",
                0,
                Map.of(),
                Map.of(),
                1,
                Map.of("EXPECT_PRESENT", 1L),
                0,
                Map.of()));

    String replayText =
        JazzerTextRenderer.renderReplay(
            Path.of("jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/seed.json"),
            outcome);

    assertTrue(
        replayText.contains(
            "Input: jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/seed.json"));
    assertFalse(replayText.contains("/Users/erst/Tools/GridGrind"));
    assertTrue(replayText.contains("Assertion Count: 1"));
    assertTrue(replayText.contains("Assertion Kinds: EXPECT_PRESENT=1"));
  }
}
