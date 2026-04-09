package dev.erst.gridgrind.jazzer.support;

import java.nio.file.Path;
import java.util.Arrays;
import java.util.Objects;

/** Enumerates the individual Jazzer harnesses that GridGrind exposes for local fuzzing. */
public enum JazzerHarness {
  PROTOCOL_REQUEST(
      "protocol-request",
      "Protocol Request",
      "dev.erst.gridgrind.jazzer.protocol.ProtocolRequestFuzzTest",
      "readRequest"),
  PROTOCOL_WORKFLOW(
      "protocol-workflow",
      "Protocol Workflow",
      "dev.erst.gridgrind.jazzer.protocol.OperationWorkflowFuzzTest",
      "executeWorkflow"),
  ENGINE_COMMAND_SEQUENCE(
      "engine-command-sequence",
      "Engine Command Sequence",
      "dev.erst.gridgrind.jazzer.engine.WorkbookCommandSequenceFuzzTest",
      "applyCommands"),
  XLSX_ROUND_TRIP(
      "xlsx-roundtrip",
      "XLSX Round Trip",
      "dev.erst.gridgrind.jazzer.engine.XlsxRoundTripFuzzTest",
      "roundTrip");

  private final String key;
  private final String displayName;
  private final String className;
  private final String methodName;

  JazzerHarness(String key, String displayName, String className, String methodName) {
    this.key = Objects.requireNonNull(key, "key must not be null");
    this.displayName = Objects.requireNonNull(displayName, "displayName must not be null");
    this.className = Objects.requireNonNull(className, "className must not be null");
    this.methodName = Objects.requireNonNull(methodName, "methodName must not be null");
  }

  /** Returns the stable external key used in scripts, run directories, and reports. */
  public String key() {
    return key;
  }

  /** Returns the human-readable harness name for operator output. */
  public String displayName() {
    return displayName;
  }

  /** Returns the fully qualified fuzz test class name. */
  public String className() {
    return className;
  }

  /** Returns the fuzz test method name that owns this harness's input directory. */
  public String methodName() {
    return methodName;
  }

  /** Returns the resource directory where promoted regression inputs for this harness live. */
  public Path inputDirectory(Path projectDirectory) {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    return projectDirectory.resolve("src/fuzz/resources").resolve(inputResourceDirectory());
  }

  /** Returns the directory where this harness's promoted metadata entries live. */
  public Path promotedMetadataDirectory(Path projectDirectory) {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    return promotedMetadataRoot(projectDirectory).resolve(key);
  }

  /** Returns the resource directory suffix used by Jazzer's regression-input discovery. */
  public String inputResourceDirectory() {
    String packagePath =
        className.substring(0, className.lastIndexOf('.')).replace('.', '/');
    String simpleName = className.substring(className.lastIndexOf('.') + 1);
    return packagePath + "/" + simpleName + "Inputs/" + methodName;
  }

  /** Resolves a harness from its stable external key. */
  public static JazzerHarness fromKey(String key) {
    Objects.requireNonNull(key, "key must not be null");
    return Arrays.stream(values())
        .filter(harness -> harness.key.equals(key))
        .findFirst()
        .orElseThrow(() -> new IllegalArgumentException("Unknown Jazzer harness: " + key));
  }

  /** Returns the project-relative root directory that owns all promoted metadata entries. */
  public static Path promotedMetadataRoot(Path projectDirectory) {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    return projectDirectory.resolve("src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata");
  }
}
