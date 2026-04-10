package dev.erst.gridgrind.jazzer.tool;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Captures the run-level metrics emitted by a completed Jazzer command. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = RunMetrics.ActiveFuzzMetrics.class, name = "ACTIVE_FUZZ"),
  @JsonSubTypes.Type(value = RunMetrics.RegressionMetrics.class, name = "REGRESSION")
})
public sealed interface RunMetrics
    permits RunMetrics.ActiveFuzzMetrics, RunMetrics.RegressionMetrics {
  /** Captures libFuzzer-style metrics from an active fuzzing session. */
  record ActiveFuzzMetrics(
      long executions,
      int coverage,
      int features,
      int corpusEntries,
      long corpusBytes,
      int maxInputBytes,
      int executionsPerSecond,
      int rssMegabytes)
      implements RunMetrics {}

  /** Captures the number of harnesses exercised during a regression replay run. */
  record RegressionMetrics(int executedHarnessCount) implements RunMetrics {}
}
