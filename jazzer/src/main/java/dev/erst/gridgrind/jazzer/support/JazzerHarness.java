package dev.erst.gridgrind.jazzer.support;

import java.nio.file.Path;
import java.util.Map;
import java.util.Objects;

/** Describes one individual Jazzer harness that GridGrind exposes for local fuzzing. */
public record JazzerHarness(String key, String displayName, String className, String methodName) {
  public JazzerHarness {
    key = requireNonBlank(key, "key");
    displayName = requireNonBlank(displayName, "displayName");
    className = requireNonBlank(className, "className");
    methodName = requireNonBlank(methodName, "methodName");
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
    String packagePath = className.substring(0, className.lastIndexOf('.')).replace('.', '/');
    String simpleName = className.substring(className.lastIndexOf('.') + 1);
    return packagePath + "/" + simpleName + "Inputs/" + methodName;
  }

  /** Returns all committed Jazzer harnesses in stable encounter order. */
  public static JazzerHarness[] values() {
    return JazzerTopology.registry().harnesses().toArray(JazzerHarness[]::new);
  }

  /** Resolves a harness from its stable external key. */
  public static JazzerHarness fromKey(String key) {
    Objects.requireNonNull(key, "key must not be null");
    Map<String, JazzerHarness> harnessesByKey = JazzerTopology.registry().harnessesByKey();
    JazzerHarness harness = harnessesByKey.get(key);
    if (harness == null) {
      throw new IllegalArgumentException("Unknown Jazzer harness: " + key);
    }
    return harness;
  }

  /** Returns the project-relative root directory that owns all promoted metadata entries. */
  public static Path promotedMetadataRoot(Path projectDirectory) {
    Objects.requireNonNull(projectDirectory, "projectDirectory must not be null");
    return projectDirectory.resolve(
        "src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata");
  }

  /** Returns the canonical protocol-request harness. */
  public static JazzerHarness protocolRequest() {
    return fromKey("protocol-request");
  }

  /** Returns the canonical protocol-workflow harness. */
  public static JazzerHarness protocolWorkflow() {
    return fromKey("protocol-workflow");
  }

  /** Returns the canonical engine-command-sequence harness. */
  public static JazzerHarness engineCommandSequence() {
    return fromKey("engine-command-sequence");
  }

  /** Returns the canonical xlsx-roundtrip harness. */
  public static JazzerHarness xlsxRoundTrip() {
    return fromKey("xlsx-roundtrip");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    String trimmed = value.trim();
    if (trimmed.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return trimmed;
  }
}
