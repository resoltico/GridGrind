package dev.erst.gridgrind.jazzer.support;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.SerializationFeature;

/** Records per-harness semantic telemetry and writes local JSON snapshots during Jazzer runs. */
public final class HarnessTelemetry {
  private static final int SNAPSHOT_INTERVAL = 250;
  private static final JsonMapper JSON_MAPPER =
      JsonMapper.builder().enable(SerializationFeature.INDENT_OUTPUT).build();
  private static final Map<JazzerHarness, HarnessTelemetry> REGISTRY = new ConcurrentHashMap<>();

  private final JazzerHarness harness;
  private final List<Path> outputDirectories;
  private final String startedAt;
  private long iterations;
  private long totalInputBytes;
  private long maxInputBytes;
  private long generatedInvalidInputs;
  private long expectedInvalidOutcomes;
  private long successfulOutcomes;
  private long unexpectedFailures;
  private Instant lastUpdatedAt;
  private final Map<String, Long> sequenceKinds = new LinkedHashMap<>();
  private final Map<String, Long> readKinds = new LinkedHashMap<>();
  private final Map<String, Long> styleKinds = new LinkedHashMap<>();
  private final Map<String, Long> sourceKinds = new LinkedHashMap<>();
  private final Map<String, Long> persistenceKinds = new LinkedHashMap<>();
  private final Map<String, Long> errorKinds = new LinkedHashMap<>();
  private final Map<String, Long> responseKinds = new LinkedHashMap<>();

  private HarnessTelemetry(JazzerHarness harness) {
    this.harness = Objects.requireNonNull(harness, "harness must not be null");
    this.outputDirectories = List.copyOf(resolveOutputDirectories());
    this.startedAt = Instant.now().toString();
    this.lastUpdatedAt = Instant.parse(startedAt);
    Runtime.getRuntime().addShutdownHook(new Thread(this::writeSnapshotSafely));
  }

  /** Returns the process-wide recorder for a single harness. */
  public static HarnessTelemetry forHarness(JazzerHarness harness) {
    Objects.requireNonNull(harness, "harness must not be null");
    return REGISTRY.computeIfAbsent(harness, HarnessTelemetry::new);
  }

  /** Starts a new fuzz iteration and records the input size in bytes. */
  public synchronized void beginIteration(int inputBytes) {
    iterations++;
    totalInputBytes += Math.max(0, inputBytes);
    maxInputBytes = Math.max(maxInputBytes, Math.max(0, inputBytes));
    touch();
    maybeWriteSnapshot();
  }

  /** Records the generated operation or command mix for the current fuzz case. */
  public synchronized void recordSequenceKinds(Map<String, Long> kinds) {
    Objects.requireNonNull(kinds, "kinds must not be null");
    kinds.forEach((key, value) -> sequenceKinds.merge(key, value, Long::sum));
    touch();
  }

  /** Records the workbook read mix observed for the current fuzz case. */
  public synchronized void recordReadKinds(Map<String, Long> kinds) {
    Objects.requireNonNull(kinds, "kinds must not be null");
    kinds.forEach((key, value) -> readKinds.merge(key, value, Long::sum));
    touch();
  }

  /** Records the style-attribute mix observed for the current fuzz case. */
  public synchronized void recordStyleKinds(Map<String, Long> kinds) {
    Objects.requireNonNull(kinds, "kinds must not be null");
    kinds.forEach((key, value) -> styleKinds.merge(key, value, Long::sum));
    touch();
  }

  /** Records the workbook source type observed for the current fuzz case. */
  public synchronized void recordSourceKind(String sourceKind) {
    Objects.requireNonNull(sourceKind, "sourceKind must not be null");
    sourceKinds.merge(sourceKind, 1L, Long::sum);
    touch();
  }

  /** Records the workbook persistence type observed for the current fuzz case. */
  public synchronized void recordPersistenceKind(String persistenceKind) {
    Objects.requireNonNull(persistenceKind, "persistenceKind must not be null");
    persistenceKinds.merge(persistenceKind, 1L, Long::sum);
    touch();
  }

  /** Records the protocol response family observed during a workflow replay or fuzz iteration. */
  public synchronized void recordResponseKind(String responseKind) {
    Objects.requireNonNull(responseKind, "responseKind must not be null");
    responseKinds.merge(responseKind, 1L, Long::sum);
    touch();
  }

  /** Records a generation-time invalid input that never reached the execution layer. */
  public synchronized void recordGeneratedInvalid() {
    generatedInvalidInputs++;
    touch();
    maybeWriteSnapshot();
  }

  /** Records a fuzz case that completed successfully. */
  public synchronized void recordSuccess() {
    successfulOutcomes++;
    touch();
    maybeWriteSnapshot();
  }

  /** Records an expected invalid execution outcome. */
  public synchronized void recordExpectedInvalid(Throwable error) {
    expectedInvalidOutcomes++;
    recordErrorKind(error);
    touch();
    maybeWriteSnapshot();
  }

  /** Records an unexpected failure that should surface as a Jazzer finding. */
  public synchronized void recordUnexpectedFailure(Throwable error) {
    unexpectedFailures++;
    recordErrorKind(error);
    touch();
    writeSnapshotSafely();
  }

  /** Returns the current immutable telemetry snapshot. */
  public synchronized HarnessTelemetrySnapshot snapshot() {
    return new HarnessTelemetrySnapshot(
        harness.key(),
        harness.displayName(),
        startedAt,
        lastUpdatedAt.toString(),
        iterations,
        totalInputBytes,
        maxInputBytes,
        generatedInvalidInputs,
        expectedInvalidOutcomes,
        successfulOutcomes,
        unexpectedFailures,
        Map.copyOf(sequenceKinds),
        Map.copyOf(readKinds),
        Map.copyOf(styleKinds),
        Map.copyOf(sourceKinds),
        Map.copyOf(persistenceKinds),
        Map.of(
            "generation_invalid", generatedInvalidInputs,
            "expected_invalid", expectedInvalidOutcomes,
            "success", successfulOutcomes,
            "unexpected_failure", unexpectedFailures),
        Map.copyOf(errorKinds),
        Map.copyOf(responseKinds));
  }

  private static List<Path> resolveOutputDirectories() {
    ArrayList<Path> directories = new ArrayList<>();
    Path workingDirectory = Path.of(System.getProperty("user.dir")).toAbsolutePath().normalize();
    directories.add(workingDirectory.resolve("telemetry"));

    String historyDirectory = System.getenv("GRIDGRIND_JAZZER_HISTORY_DIR");
    if (historyDirectory != null && !historyDirectory.isBlank()) {
      Path historyPath = Path.of(historyDirectory).toAbsolutePath().normalize();
      Path telemetryDirectory = historyPath.resolve("telemetry");
      if (!directories.contains(telemetryDirectory)) {
        directories.add(telemetryDirectory);
      }
    }
    return directories;
  }

  private void maybeWriteSnapshot() {
    if (iterations % SNAPSHOT_INTERVAL == 0) {
      writeSnapshotSafely();
    }
  }

  private synchronized void touch() {
    lastUpdatedAt = Instant.now();
  }

  private synchronized void recordErrorKind(Throwable error) {
    Objects.requireNonNull(error, "error must not be null");
    errorKinds.merge(error.getClass().getSimpleName(), 1L, Long::sum);
  }

  private void writeSnapshotSafely() {
    try {
      writeSnapshot();
    } catch (IOException ignored) {
      // Telemetry is observability only and must never hide the actual fuzzing result.
    }
  }

  private synchronized void writeSnapshot() throws IOException {
    HarnessTelemetrySnapshot snapshot = snapshot();
    for (Path directory : outputDirectories) {
      Files.createDirectories(directory);
      Path snapshotPath = directory.resolve(harness.key() + ".json");
      JSON_MAPPER.writeValue(snapshotPath.toFile(), snapshot);
    }
  }
}
