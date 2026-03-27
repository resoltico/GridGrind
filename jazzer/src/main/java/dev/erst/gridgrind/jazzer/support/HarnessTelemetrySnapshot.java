package dev.erst.gridgrind.jazzer.support;

import java.util.Map;

/** Captures the semantic run metrics that a single Jazzer harness recorded during one JVM run. */
public record HarnessTelemetrySnapshot(
    String harnessKey,
    String displayName,
    String startedAt,
    String updatedAt,
    long iterations,
    long totalInputBytes,
    long maxInputBytes,
    long generatedInvalidInputs,
    long expectedInvalidOutcomes,
    long successfulOutcomes,
    long unexpectedFailures,
    Map<String, Long> sequenceKinds,
    Map<String, Long> readKinds,
    Map<String, Long> styleKinds,
    Map<String, Long> sourceKinds,
    Map<String, Long> persistenceKinds,
    Map<String, Long> outcomeKinds,
    Map<String, Long> errorKinds,
    Map<String, Long> responseKinds) {

  /** Normalizes deserialized telemetry so older local history remains readable as the schema grows. */
  public HarnessTelemetrySnapshot {
    sequenceKinds = Map.copyOf(sequenceKinds == null ? Map.of() : sequenceKinds);
    readKinds = Map.copyOf(readKinds == null ? Map.of() : readKinds);
    styleKinds = Map.copyOf(styleKinds == null ? Map.of() : styleKinds);
    sourceKinds = Map.copyOf(sourceKinds == null ? Map.of() : sourceKinds);
    persistenceKinds = Map.copyOf(persistenceKinds == null ? Map.of() : persistenceKinds);
    outcomeKinds = Map.copyOf(outcomeKinds == null ? Map.of() : outcomeKinds);
    errorKinds = Map.copyOf(errorKinds == null ? Map.of() : errorKinds);
    responseKinds = Map.copyOf(responseKinds == null ? Map.of() : responseKinds);
  }
}
