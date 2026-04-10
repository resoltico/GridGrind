package dev.erst.gridgrind.jazzer.tool;

import dev.erst.gridgrind.jazzer.support.HarnessTelemetrySnapshot;
import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.JazzerRunTarget;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

/** Builds, persists, and reads the local summary artifacts that surround Jazzer runs. */
public final class JazzerReportSupport {
  private static final Pattern DONE_PATTERN =
      Pattern.compile(
          "#\\d+\\s+DONE\\s+cov:\\s+(\\d+)\\s+ft:\\s+(\\d+)\\s+corp:\\s+(\\d+)/(\\d+)([KMG]?b)\\s+lim:\\s+(\\d+)\\s+exec/s:\\s+(\\d+)\\s+rss:\\s+(\\d+)Mb");
  private static final Pattern RUNS_PATTERN =
      Pattern.compile("Done\\s+(\\d+)\\s+runs\\s+in\\s+(\\d+)\\s+second\\(s\\)");
  private static final List<String> RAW_FINDING_PREFIXES =
      List.of("crash-", "timeout-", "oom-", "leak-");

  private JazzerReportSupport() {}

  /** Builds one completed run summary and writes its JSON and text artifacts. */
  public static LocalRunSummary summarizeRun(
      Path projectDirectory,
      JazzerRunTarget target,
      String taskName,
      Path logPath,
      Path historyDirectory,
      String startedAt,
      String finishedAt,
      int exitCode,
      CorpusStats corpusBefore)
      throws IOException {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(taskName, "taskName must not be null");
    Objects.requireNonNull(logPath, "logPath must not be null");
    Objects.requireNonNull(historyDirectory, "historyDirectory must not be null");
    Objects.requireNonNull(startedAt, "startedAt must not be null");
    Objects.requireNonNull(finishedAt, "finishedAt must not be null");
    Objects.requireNonNull(corpusBefore, "corpusBefore must not be null");

    Path runDirectory = target.workingDirectory(projectDirectory);
    CorpusStats corpusAfter = scanCorpus(runDirectory);
    List<HarnessTelemetrySnapshot> telemetry =
        readTelemetry(runDirectory, historyDirectory, target);
    List<FindingArtifact> findings = writeFindingArtifacts(runDirectory, historyDirectory, target);
    RunMetrics metrics = parseMetrics(logPath, telemetry, target);
    LocalRunSummary summary =
        new LocalRunSummary(
            target.key(),
            target.displayName(),
            taskName,
            target.activeFuzzing() ? RunMode.ACTIVE_FUZZING : RunMode.REGRESSION,
            startedAt,
            finishedAt,
            Duration.between(Instant.parse(startedAt), Instant.parse(finishedAt)).toSeconds(),
            exitCode,
            exitCode == 0 ? "SUCCESS" : "FAILURE",
            logPath.toAbsolutePath().normalize().toString(),
            historyDirectory.toAbsolutePath().normalize().toString(),
            corpusBefore,
            corpusAfter,
            metrics,
            List.copyOf(telemetry),
            List.copyOf(findings));
    writeSummaryArtifacts(runDirectory, historyDirectory, summary);
    return summary;
  }

  /** Returns the latest summary if the target has already been run locally. */
  public static LocalRunSummary readLatestSummary(Path projectDirectory, JazzerRunTarget target)
      throws IOException {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Objects.requireNonNull(target, "target must not be null");
    return JazzerJson.read(
        latestSummaryJson(target.workingDirectory(projectDirectory)), LocalRunSummary.class);
  }

  /** Returns whether the target already has a latest-summary artifact on disk. */
  public static boolean hasLatestSummary(Path projectDirectory, JazzerRunTarget target) {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Objects.requireNonNull(target, "target must not be null");
    return Files.exists(latestSummaryJson(target.workingDirectory(projectDirectory)));
  }

  /** Scans one target's local corpus. */
  public static CorpusStats scanCorpus(Path runDirectory) throws IOException {
    Objects.requireNonNull(runDirectory, "runDirectory must not be null");
    Path corpusDirectory = runDirectory.resolve(".cifuzz-corpus");
    if (!Files.exists(corpusDirectory)) {
      return new CorpusStats(0L, 0L);
    }
    try (Stream<Path> stream = Files.walk(corpusDirectory)) {
      long fileCount = stream.filter(Files::isRegularFile).count();
      try (Stream<Path> sizeStream = Files.walk(corpusDirectory)) {
        long totalBytes =
            sizeStream
                .filter(Files::isRegularFile)
                .mapToLong(
                    path -> {
                      try {
                        return Files.size(path);
                      } catch (IOException exception) {
                        throw new IllegalStateException(
                            "Failed to read corpus file size: " + path, exception);
                      }
                    })
                .sum();
        return new CorpusStats(fileCount, totalBytes);
      }
    }
  }

  /** Returns the newest raw corpus entries for one run directory. */
  public static List<Path> newestCorpusEntries(Path runDirectory, int limit) throws IOException {
    Objects.requireNonNull(runDirectory, "runDirectory must not be null");
    Path corpusDirectory = runDirectory.resolve(".cifuzz-corpus");
    if (!Files.exists(corpusDirectory)) {
      return List.of();
    }
    try (Stream<Path> stream = Files.walk(corpusDirectory)) {
      return stream
          .filter(Files::isRegularFile)
          .sorted(Comparator.comparingLong(JazzerReportSupport::lastModifiedMillis).reversed())
          .limit(limit)
          .toList();
    }
  }

  /** Returns the promoted regression inputs committed for one harness. */
  public static List<Path> promotedInputs(Path projectDirectory, JazzerHarness harness)
      throws IOException {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Objects.requireNonNull(harness, "harness must not be null");
    Path inputDirectory = harness.inputDirectory(projectDirectory);
    if (!Files.exists(inputDirectory)) {
      return List.of();
    }
    try (Stream<Path> stream = Files.list(inputDirectory)) {
      return stream.filter(Files::isRegularFile).sorted().toList();
    }
  }

  /**
   * Returns input files in the harness input directory that have no corresponding promoted-metadata
   * entry. An orphaned input was committed directly to the input directory without going through
   * {@code jazzer/bin/promote}, leaving it without a replay contract enforced by {@code
   * PromotionMetadataTest}.
   */
  public static List<Path> orphanedInputs(Path projectDirectory, JazzerHarness harness)
      throws IOException {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Objects.requireNonNull(harness, "harness must not be null");
    List<Path> inputs = promotedInputs(projectDirectory, harness);
    if (inputs.isEmpty()) {
      return List.of();
    }
    Set<Path> promotedPaths = promotedInputPaths(projectDirectory);
    return inputs.stream()
        .filter(input -> !promotedPaths.contains(input.toAbsolutePath().normalize()))
        .sorted()
        .toList();
  }

  /**
   * Returns the set of {@code promotedInputPath} values recorded across all promoted-metadata
   * entries in the project, normalized to absolute paths.
   */
  public static Set<Path> promotedInputPaths(Path projectDirectory) throws IOException {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    Path metadataRoot = JazzerHarness.promotedMetadataRoot(projectDirectory);
    if (!Files.exists(metadataRoot)) {
      return Set.of();
    }
    Set<Path> result = new HashSet<>();
    try (Stream<Path> stream = Files.walk(metadataRoot)) {
      for (Path jsonPath :
          stream.filter(p -> p.getFileName().toString().endsWith(".json")).sorted().toList()) {
        PromotionMetadata metadata = JazzerJson.read(jsonPath, PromotionMetadata.class);
        result.add(metadata.promotedInputPath(projectDirectory));
      }
    }
    return Set.copyOf(result);
  }

  /** Returns file-count and byte-count statistics for an arbitrary list of input files. */
  public static CorpusStats scanFiles(List<Path> files) throws IOException {
    Objects.requireNonNull(files, "files must not be null");

    long totalBytes = 0L;
    for (Path path : files) {
      totalBytes += Files.size(path);
    }
    return new CorpusStats(files.size(), totalBytes);
  }

  /** Returns the replayed finding artifacts currently recorded for one run directory. */
  public static List<FindingArtifact> findingArtifacts(Path runDirectory) throws IOException {
    Objects.requireNonNull(runDirectory, "runDirectory must not be null");
    Path findingsDirectory = runDirectory.resolve("findings");
    if (!Files.exists(findingsDirectory)) {
      return List.of();
    }

    ArrayList<FindingArtifact> findings = new ArrayList<>();
    try (Stream<Path> stream = Files.list(findingsDirectory)) {
      for (Path jsonPath :
          stream
              .filter(path -> path.getFileName().toString().endsWith(".json"))
              .sorted()
              .toList()) {
        findings.add(JazzerJson.read(jsonPath, FindingArtifact.class));
      }
    }
    return List.copyOf(findings);
  }

  private static RunMetrics parseMetrics(
      Path logPath, List<HarnessTelemetrySnapshot> telemetry, JazzerRunTarget target)
      throws IOException {
    if (!target.activeFuzzing()) {
      return new RunMetrics.RegressionMetrics(telemetry.size());
    }

    return parseActiveFuzzMetrics(Files.readString(logPath));
  }

  /** Parses the libFuzzer summary block from one active-fuzz run log. */
  static RunMetrics.ActiveFuzzMetrics parseActiveFuzzMetrics(String log) {
    Matcher runsMatcher = RUNS_PATTERN.matcher(log);
    Matcher doneMatcher = DONE_PATTERN.matcher(log);
    if (!runsMatcher.find() || !doneMatcher.find()) {
      return new RunMetrics.ActiveFuzzMetrics(0L, 0, 0, 0, 0L, 0, 0, 0);
    }
    return new RunMetrics.ActiveFuzzMetrics(
        Long.parseLong(runsMatcher.group(1)),
        Integer.parseInt(doneMatcher.group(1)),
        Integer.parseInt(doneMatcher.group(2)),
        Integer.parseInt(doneMatcher.group(3)),
        parseCorpusBytes(doneMatcher.group(4), doneMatcher.group(5)),
        Integer.parseInt(doneMatcher.group(6)),
        Integer.parseInt(doneMatcher.group(7)),
        Integer.parseInt(doneMatcher.group(8)));
  }

  private static long parseCorpusBytes(String value, String unit) {
    long magnitude = Long.parseLong(value);
    return switch (unit) {
      case "b" -> magnitude;
      case "Kb" -> magnitude * 1024L;
      case "Mb" -> magnitude * 1024L * 1024L;
      case "Gb" -> magnitude * 1024L * 1024L * 1024L;
      default -> throw new IllegalArgumentException("Unsupported corpus byte unit: " + unit);
    };
  }

  private static List<HarnessTelemetrySnapshot> readTelemetry(
      Path runDirectory, Path historyDirectory, JazzerRunTarget target) throws IOException {
    ArrayList<HarnessTelemetrySnapshot> snapshots = new ArrayList<>();
    Path historyTelemetryDirectory = historyDirectory.resolve("telemetry");
    Path runTelemetryDirectory = runDirectory.resolve("telemetry");
    for (JazzerHarness harness : target.harnesses()) {
      Path historyTelemetryPath = historyTelemetryDirectory.resolve(harness.key() + ".json");
      Path runTelemetryPath = runTelemetryDirectory.resolve(harness.key() + ".json");
      Path sourcePath =
          Files.exists(historyTelemetryPath) ? historyTelemetryPath : runTelemetryPath;
      if (Files.exists(sourcePath)) {
        snapshots.add(JazzerJson.read(sourcePath, HarnessTelemetrySnapshot.class));
      }
    }
    return List.copyOf(snapshots);
  }

  private static List<FindingArtifact> writeFindingArtifacts(
      Path runDirectory, Path historyDirectory, JazzerRunTarget target) throws IOException {
    if (!target.replayable()) {
      return List.copyOf(findingArtifacts(runDirectory));
    }

    ArrayList<FindingArtifact> findings = new ArrayList<>();
    for (Path rawArtifact : rawFindingArtifacts(runDirectory)) {
      ReplayOutcome replayOutcome =
          JazzerReplaySupport.replay(target.replayHarness(), Files.readAllBytes(rawArtifact));
      FindingArtifact finding =
          new FindingArtifact(
              rawArtifact.getFileName().toString(),
              rawArtifact.toAbsolutePath().normalize().toString(),
              replayOutcomeKind(replayOutcome),
              runDirectory
                  .resolve("findings")
                  .resolve(rawArtifact.getFileName() + ".json")
                  .toString(),
              runDirectory
                  .resolve("findings")
                  .resolve(rawArtifact.getFileName() + ".txt")
                  .toString());
      writeFindingArtifact(
          runDirectory.resolve("findings"),
          rawArtifact.getFileName().toString(),
          finding,
          replayOutcome);
      writeFindingArtifact(
          historyDirectory.resolve("findings"),
          rawArtifact.getFileName().toString(),
          finding,
          replayOutcome);
      findings.add(finding);
    }
    return List.copyOf(findings);
  }

  private static void writeFindingArtifact(
      Path outputDirectory, String baseName, FindingArtifact finding, ReplayOutcome replayOutcome)
      throws IOException {
    Files.createDirectories(outputDirectory);
    JazzerJson.write(outputDirectory.resolve(baseName + ".json"), finding);
    Files.writeString(
        outputDirectory.resolve(baseName + ".txt"),
        JazzerTextRenderer.renderReplay(Path.of(finding.rawArtifactPath()), replayOutcome)
            + System.lineSeparator());
  }

  private static List<Path> rawFindingArtifacts(Path runDirectory) throws IOException {
    if (!Files.exists(runDirectory)) {
      return List.of();
    }
    try (Stream<Path> stream = Files.list(runDirectory)) {
      return stream
          .filter(Files::isRegularFile)
          .filter(
              path ->
                  RAW_FINDING_PREFIXES.stream()
                      .anyMatch(prefix -> path.getFileName().toString().startsWith(prefix)))
          .sorted()
          .toList();
    }
  }

  private static void writeSummaryArtifacts(
      Path runDirectory, Path historyDirectory, LocalRunSummary summary) throws IOException {
    Files.createDirectories(historyDirectory);
    JazzerJson.write(latestSummaryJson(runDirectory), summary);
    Files.writeString(
        latestSummaryText(runDirectory),
        JazzerTextRenderer.renderSummary(summary) + System.lineSeparator());
    JazzerJson.write(historyDirectory.resolve("summary.json"), summary);
    Files.writeString(
        historyDirectory.resolve("summary.txt"),
        JazzerTextRenderer.renderSummary(summary) + System.lineSeparator());
  }

  private static Path latestSummaryJson(Path runDirectory) {
    return runDirectory.resolve("latest-summary.json");
  }

  private static Path latestSummaryText(Path runDirectory) {
    return runDirectory.resolve("latest-summary.txt");
  }

  private static String replayOutcomeKind(ReplayOutcome outcome) {
    return switch (outcome) {
      case ReplayOutcome.Success _ -> "SUCCESS";
      case ReplayOutcome.ExpectedInvalid _ -> "EXPECTED_INVALID";
      case ReplayOutcome.UnexpectedFailure _ -> "UNEXPECTED_FAILURE";
    };
  }

  private static long lastModifiedMillis(Path path) {
    try {
      return Files.getLastModifiedTime(path).toMillis();
    } catch (IOException exception) {
      throw new IllegalStateException("Failed to read last-modified time for " + path, exception);
    }
  }
}
