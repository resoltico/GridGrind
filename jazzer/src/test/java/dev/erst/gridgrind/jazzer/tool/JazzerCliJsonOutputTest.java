package dev.erst.gridgrind.jazzer.tool;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.JazzerRunTarget;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;

/** Locks the machine-readable Jazzer operator surface for agent and shell consumers. */
class JazzerCliJsonOutputTest {
  private static final JsonMapper JSON_MAPPER = JsonMapper.builder().build();

  @Test
  void statusAndReportEmitSummaryJson(@TempDir Path projectDirectory) throws IOException {
    Path runDirectory = JazzerRunTarget.protocolRequest().workingDirectory(projectDirectory);
    Files.createDirectories(runDirectory);
    JazzerJson.write(runDirectory.resolve("latest-summary.json"), sampleSummary());

    JsonNode statusJson =
        runJson(projectDirectory, "status", "--target", "protocol-request", "--json");
    JsonNode reportJson =
        runJson(projectDirectory, "report", "--target", "protocol-request", "--json");

    assertEquals(
        "protocol-request", statusJson.path("summaries").get(0).path("targetKey").stringValue());
    assertEquals(
        "protocol-request", reportJson.path("summaries").get(0).path("targetKey").stringValue());
    assertEquals(
        "jazzer/.local/runs/protocol-request/log.txt",
        statusJson.path("summaries").get(0).path("logPath").stringValue());
    assertEquals(
        "jazzer/.local/runs/protocol-request/log.txt",
        reportJson.path("summaries").get(0).path("logPath").stringValue());
  }

  @Test
  void listFindingsEmitsTargetJson(@TempDir Path projectDirectory) throws IOException {
    Path findingsDirectory =
        JazzerRunTarget.protocolRequest().workingDirectory(projectDirectory).resolve("findings");
    Files.createDirectories(findingsDirectory);
    JazzerJson.write(
        findingsDirectory.resolve("crash-clean.json"),
        new FindingArtifact(
            "crash-clean",
            "/tmp/crash-clean",
            "SUCCESS",
            "/tmp/crash-clean.json",
            "/tmp/crash-clean.txt"));

    JsonNode findingsJson =
        runJson(projectDirectory, "list-findings", "--target", "protocol-request", "--json");

    JsonNode targetJson = findingsJson.path("targets").get(0);
    assertEquals("protocol-request", targetJson.path("target").stringValue());
    assertEquals(
        "crash-clean", targetJson.path("findings").get(0).path("rawArtifactName").stringValue());
    assertEquals("SUCCESS", targetJson.path("findings").get(0).path("replayOutcome").stringValue());
  }

  @Test
  void listCorpusEmitsCorpusAndPromotionJson(@TempDir Path projectDirectory) throws IOException {
    Path runDirectory = JazzerRunTarget.protocolRequest().workingDirectory(projectDirectory);
    Path corpusEntry = runDirectory.resolve(".cifuzz-corpus/seed.bin");
    Files.createDirectories(corpusEntry.getParent());
    Files.write(corpusEntry, new byte[] {1, 2, 3});

    JazzerHarness harness = JazzerHarness.protocolRequest();
    Path promotedInputPath = harness.inputDirectory(projectDirectory).resolve("seed.json");
    Files.createDirectories(promotedInputPath.getParent());
    Files.writeString(promotedInputPath, "{}");

    Path metadataPath = harness.promotedMetadataDirectory(projectDirectory).resolve("seed.json");
    Files.createDirectories(metadataPath.getParent());
    Path replayTextPath = metadataPath.resolveSibling("seed.txt");
    Files.writeString(replayTextPath, "Replay Result" + System.lineSeparator());
    JazzerJson.write(
        metadataPath,
        new PromotionMetadata(
            harness.key(),
            PromotionMetadata.relativizePath(projectDirectory, promotedInputPath),
            PromotionMetadata.relativizePath(projectDirectory, promotedInputPath),
            "SUCCESS",
            new ReplayExpectation(
                "SUCCESS",
                new ProtocolRequestDetails(
                    2, "PARSED", "NEW", "NONE", 0, Map.of(), Map.of(), 0, Map.of(), 0, Map.of())),
            "2026-05-26T00:00:00Z",
            PromotionMetadata.relativizePath(projectDirectory, replayTextPath)));

    JsonNode corpusJson =
        runJson(projectDirectory, "list-corpus", "--target", "protocol-request", "--json");

    JsonNode targetJson = corpusJson.path("targets").get(0);
    assertEquals("protocol-request", targetJson.path("target").stringValue());
    assertEquals(1, targetJson.path("generatedLocalCorpus").path("fileCount").asInt());
    assertEquals(3L, targetJson.path("generatedLocalCorpus").path("totalBytes").asLong());
    assertEquals(1, targetJson.path("committedCustomSeeds").path("fileCount").asInt());
    assertEquals(2L, targetJson.path("committedCustomSeeds").path("totalBytes").asLong());
    assertEquals(1, targetJson.path("newestLocalCorpusEntries").size());
    assertEquals(1, targetJson.path("promotedInputs").size());
    assertEquals(0, targetJson.path("orphanedInputs").size());
    assertTrue(
        targetJson.path("newestLocalCorpusEntries").get(0).stringValue().endsWith("seed.bin"));
    assertTrue(targetJson.path("promotedInputs").get(0).stringValue().endsWith("seed.json"));
  }

  private static JsonNode runJson(Path projectDirectory, String command, String... extraArguments)
      throws IOException {
    String[] arguments = new String[extraArguments.length + 3];
    arguments[0] = command;
    arguments[1] = "--project-dir";
    arguments[2] = projectDirectory.toString();
    System.arraycopy(extraArguments, 0, arguments, 3, extraArguments.length);

    ByteArrayOutputStream standardOutputBuffer = new ByteArrayOutputStream();
    ByteArrayOutputStream errorOutputBuffer = new ByteArrayOutputStream();
    try (PrintStream standardOutput =
            new PrintStream(standardOutputBuffer, true, StandardCharsets.UTF_8);
        PrintStream errorOutput =
            new PrintStream(errorOutputBuffer, true, StandardCharsets.UTF_8)) {
      int exitCode = JazzerCli.run(arguments, standardOutput, errorOutput);
      assertEquals(0, exitCode);
    }
    assertEquals("", errorOutputBuffer.toString(StandardCharsets.UTF_8));
    return JSON_MAPPER.readTree(standardOutputBuffer.toString(StandardCharsets.UTF_8));
  }

  private static LocalRunSummary sampleSummary() {
    return new LocalRunSummary(
        "protocol-request",
        "Protocol Request",
        "fuzzProtocolRequest",
        RunMode.REGRESSION,
        "2026-05-26T00:00:00Z",
        "2026-05-26T00:00:01Z",
        1L,
        0,
        "SUCCESS",
        "jazzer/.local/runs/protocol-request/log.txt",
        "jazzer/.local/runs/protocol-request/history",
        new CorpusStats(0L, 0L),
        new CorpusStats(1L, 3L),
        new RunMetrics.RegressionMetrics(1),
        List.of(),
        List.of());
  }
}
