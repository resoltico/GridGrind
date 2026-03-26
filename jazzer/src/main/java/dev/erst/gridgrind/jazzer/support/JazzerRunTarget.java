package dev.erst.gridgrind.jazzer.support;

import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/** Enumerates the runnable Jazzer targets, including the aggregate regression target. */
public enum JazzerRunTarget {
  REGRESSION(
      "regression",
      "Regression Replay",
      "jazzerRegression",
      ".local/runs/regression",
      false,
      List.of(
          JazzerHarness.PROTOCOL_REQUEST,
          JazzerHarness.PROTOCOL_WORKFLOW,
          JazzerHarness.ENGINE_COMMAND_SEQUENCE,
          JazzerHarness.XLSX_ROUND_TRIP)),
  PROTOCOL_REQUEST(
      "protocol-request",
      "Protocol Request",
      "fuzzProtocolRequest",
      ".local/runs/protocol-request",
      true,
      List.of(JazzerHarness.PROTOCOL_REQUEST)),
  PROTOCOL_WORKFLOW(
      "protocol-workflow",
      "Protocol Workflow",
      "fuzzProtocolWorkflow",
      ".local/runs/protocol-workflow",
      true,
      List.of(JazzerHarness.PROTOCOL_WORKFLOW)),
  ENGINE_COMMAND_SEQUENCE(
      "engine-command-sequence",
      "Engine Command Sequence",
      "fuzzEngineCommandSequence",
      ".local/runs/engine-command-sequence",
      true,
      List.of(JazzerHarness.ENGINE_COMMAND_SEQUENCE)),
  XLSX_ROUND_TRIP(
      "xlsx-roundtrip",
      "XLSX Round Trip",
      "fuzzXlsxRoundTrip",
      ".local/runs/xlsx-roundtrip",
      true,
      List.of(JazzerHarness.XLSX_ROUND_TRIP));

  private final String key;
  private final String displayName;
  private final String taskName;
  private final String workingDirectory;
  private final boolean activeFuzzing;
  private final List<JazzerHarness> harnesses;

  JazzerRunTarget(
      String key,
      String displayName,
      String taskName,
      String workingDirectory,
      boolean activeFuzzing,
      List<JazzerHarness> harnesses) {
    this.key = Objects.requireNonNull(key, "key must not be null");
    this.displayName = Objects.requireNonNull(displayName, "displayName must not be null");
    this.taskName = Objects.requireNonNull(taskName, "taskName must not be null");
    this.workingDirectory = Objects.requireNonNull(workingDirectory, "workingDirectory must not be null");
    this.activeFuzzing = activeFuzzing;
    this.harnesses = List.copyOf(harnesses);
  }

  /** Returns the stable external key used by scripts, local directories, and reports. */
  public String key() {
    return key;
  }

  /** Returns the human-readable target name for operator output. */
  public String displayName() {
    return displayName;
  }

  /** Returns the Gradle task name that launches this run target. */
  public String taskName() {
    return taskName;
  }

  /** Returns the target working directory relative to the Jazzer project root. */
  public String workingDirectory() {
    return workingDirectory;
  }

  /** Returns whether this target performs active fuzzing rather than regression replay. */
  public boolean activeFuzzing() {
    return activeFuzzing;
  }

  /** Returns the harnesses executed by this target. */
  public List<JazzerHarness> harnesses() {
    return harnesses;
  }

  /** Returns the absolute working directory for this target under the Jazzer project root. */
  public Path workingDirectory(Path projectDirectory) {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    return projectDirectory.resolve(workingDirectory);
  }

  /** Returns whether this target corresponds to exactly one replayable harness. */
  public boolean replayable() {
    return harnesses.size() == 1;
  }

  /** Returns the single harness behind this target, rejecting aggregate targets like regression. */
  public JazzerHarness replayHarness() {
    if (!replayable()) {
      throw new IllegalArgumentException("Target is not replayable: " + key);
    }
    return harnesses.getFirst();
  }

  /** Resolves a run target from its stable external key. */
  public static JazzerRunTarget fromKey(String key) {
    Objects.requireNonNull(key, "key must not be null");
    return Arrays.stream(values())
        .filter(target -> target.key.equals(key))
        .findFirst()
        .orElseThrow(() -> new IllegalArgumentException("Unknown Jazzer run target: " + key));
  }

  /** Resolves a run target from its Gradle task name. */
  public static JazzerRunTarget fromTaskName(String taskName) {
    Objects.requireNonNull(taskName, "taskName must not be null");
    return Arrays.stream(values())
        .filter(target -> target.taskName.equals(taskName))
        .findFirst()
        .orElseThrow(() -> new IllegalArgumentException("Unknown Jazzer task: " + taskName));
  }
}
