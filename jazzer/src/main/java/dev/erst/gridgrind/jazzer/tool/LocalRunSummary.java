package dev.erst.gridgrind.jazzer.tool;

import dev.erst.gridgrind.jazzer.support.HarnessTelemetrySnapshot;
import java.util.List;

/** Aggregates one completed local Jazzer command into a machine-readable run summary. */
public record LocalRunSummary(
    String targetKey,
    String displayName,
    String taskName,
    RunMode mode,
    String startedAt,
    String finishedAt,
    long durationSeconds,
    int exitCode,
    String outcome,
    String logPath,
    String historyDirectory,
    CorpusStats corpusBefore,
    CorpusStats corpusAfter,
    RunMetrics metrics,
    List<HarnessTelemetrySnapshot> harnessTelemetry,
    List<FindingArtifact> findings) {}
