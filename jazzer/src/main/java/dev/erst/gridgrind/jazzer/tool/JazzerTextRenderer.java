package dev.erst.gridgrind.jazzer.tool;

import dev.erst.gridgrind.jazzer.support.HarnessTelemetrySnapshot;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** Renders human-readable summaries for the local Jazzer operator workflow. */
public final class JazzerTextRenderer {
  private JazzerTextRenderer() {}

  /** Renders one detailed run summary for terminal output. */
  public static String renderSummary(LocalRunSummary summary) {
    Objects.requireNonNull(summary, "summary must not be null");

    StringBuilder builder = new StringBuilder();
    builder.append("Jazzer Run Summary").append(System.lineSeparator());
    builder
        .append("Target: ")
        .append(summary.displayName())
        .append(" [")
        .append(summary.targetKey())
        .append("]")
        .append(System.lineSeparator());
    builder.append("Mode: ").append(summary.mode()).append(System.lineSeparator());
    builder
        .append("Outcome: ")
        .append(summary.outcome())
        .append(" (exit ")
        .append(summary.exitCode())
        .append(")")
        .append(System.lineSeparator());
    builder.append("Started: ").append(summary.startedAt()).append(System.lineSeparator());
    builder.append("Finished: ").append(summary.finishedAt()).append(System.lineSeparator());
    builder
        .append("Duration: ")
        .append(summary.durationSeconds())
        .append("s")
        .append(System.lineSeparator());
    builder.append("Log: ").append(summary.logPath()).append(System.lineSeparator());
    builder.append("History: ").append(summary.historyDirectory()).append(System.lineSeparator());
    builder.append(System.lineSeparator());
    builder
        .append("Corpus: ")
        .append(summary.corpusBefore().fileCount())
        .append(" -> ")
        .append(summary.corpusAfter().fileCount())
        .append(" files, ")
        .append(summary.corpusBefore().totalBytes())
        .append(" -> ")
        .append(summary.corpusAfter().totalBytes())
        .append(" bytes")
        .append(System.lineSeparator());
    builder.append(renderMetrics(summary.metrics()));
    builder.append(System.lineSeparator());
    long activeFindings = actionableFindingCount(summary.findings());
    long expectedInvalidArtifacts = expectedInvalidArtifactCount(summary.findings());
    long replayCleanArtifacts = replayCleanArtifactCount(summary.findings());
    builder.append("Findings: ").append(activeFindings).append(" active");
    if (expectedInvalidArtifacts > 0) {
      builder.append(", ").append(expectedInvalidArtifacts).append(" expected-invalid artifacts");
    }
    if (replayCleanArtifacts > 0) {
      builder.append(", ").append(replayCleanArtifacts).append(" replay-clean artifacts");
    }
    builder.append(System.lineSeparator());
    for (FindingArtifact finding : summary.findings()) {
      builder
          .append("  - ")
          .append(finding.rawArtifactName())
          .append(" [")
          .append(finding.replayOutcome())
          .append("]")
          .append(System.lineSeparator());
    }
    if (!summary.harnessTelemetry().isEmpty()) {
      builder.append(System.lineSeparator());
      builder.append("Harness Telemetry").append(System.lineSeparator());
      for (HarnessTelemetrySnapshot snapshot : summary.harnessTelemetry()) {
        builder.append(renderHarnessTelemetry(snapshot));
      }
    }
    return builder.toString().trim();
  }

  /** Renders a concise cross-target status view from the latest summaries. */
  public static String renderStatus(List<LocalRunSummary> summaries) {
    Objects.requireNonNull(summaries, "summaries must not be null");
    StringBuilder builder = new StringBuilder();
    builder.append("Jazzer Status").append(System.lineSeparator());
    if (summaries.isEmpty()) {
      builder.append("(no summaries recorded)").append(System.lineSeparator());
      return builder.toString().trim();
    }
    for (LocalRunSummary summary : summaries) {
      builder
          .append("- ")
          .append(summary.targetKey())
          .append(": ")
          .append(summary.outcome())
          .append(", ")
          .append(summary.durationSeconds())
          .append("s, findings=")
          .append(actionableFindingCount(summary.findings()));
      long expectedInvalidArtifacts = expectedInvalidArtifactCount(summary.findings());
      if (expectedInvalidArtifacts > 0) {
        builder.append(", expected-invalid=").append(expectedInvalidArtifacts);
      }
      long replayCleanArtifacts = replayCleanArtifactCount(summary.findings());
      if (replayCleanArtifacts > 0) {
        builder.append(", replay-clean=").append(replayCleanArtifacts);
      }
      switch (summary.metrics()) {
        case RunMetrics.ActiveFuzzMetrics metrics ->
            builder
                .append(", exec=")
                .append(metrics.executions())
                .append(", cov=")
                .append(metrics.coverage());
        case RunMetrics.RegressionMetrics metrics ->
            builder.append(", harnesses=").append(metrics.executedHarnessCount());
      }
      builder.append(System.lineSeparator());
    }
    return builder.toString().trim();
  }

  /** Renders one replay outcome for a raw local input artifact. */
  public static String renderReplay(Path inputPath, ReplayOutcome outcome) {
    Objects.requireNonNull(inputPath, "inputPath must not be null");
    Objects.requireNonNull(outcome, "outcome must not be null");

    StringBuilder builder = new StringBuilder();
    builder.append("Replay Result").append(System.lineSeparator());
    builder.append("Input: ").append(inputPath.normalize()).append(System.lineSeparator());
    builder.append("Harness: ").append(outcome.harnessKey()).append(System.lineSeparator());
    switch (outcome) {
      case ReplayOutcome.Success success -> {
        builder.append("Outcome: SUCCESS").append(System.lineSeparator());
        builder.append(renderReplayDetails(success.details()));
      }
      case ReplayOutcome.ExpectedInvalid invalid -> {
        builder.append("Outcome: EXPECTED_INVALID").append(System.lineSeparator());
        builder.append("Kind: ").append(invalid.invalidKind()).append(System.lineSeparator());
        builder.append("Message: ").append(invalid.message()).append(System.lineSeparator());
        builder.append(renderReplayDetails(invalid.details()));
      }
      case ReplayOutcome.UnexpectedFailure failure -> {
        builder.append("Outcome: UNEXPECTED_FAILURE").append(System.lineSeparator());
        builder.append("Kind: ").append(failure.failureKind()).append(System.lineSeparator());
        builder.append("Message: ").append(failure.message()).append(System.lineSeparator());
        builder.append(renderReplayDetails(failure.details())).append(System.lineSeparator());
        builder.append("Stack Trace").append(System.lineSeparator());
        builder.append(failure.stackTrace().strip());
      }
    }
    return builder.toString().trim();
  }

  /** Renders the local-corpus and committed-seed inventory for one replayable harness target. */
  static String renderCorpusListing(
      String targetKey,
      CorpusStats localCorpusStats,
      CorpusStats committedSeedStats,
      List<Path> newestEntries,
      List<Path> promotedInputs,
      List<Path> orphanedInputs) {
    StringBuilder builder = new StringBuilder();
    builder.append("Corpus Listing").append(System.lineSeparator());
    builder.append("Target: ").append(targetKey).append(System.lineSeparator());
    builder
        .append("Generated Local Corpus: ")
        .append(localCorpusStats.fileCount())
        .append(" files, ")
        .append(localCorpusStats.totalBytes())
        .append(" bytes")
        .append(System.lineSeparator());
    builder
        .append("Committed Custom Seeds: ")
        .append(committedSeedStats.fileCount())
        .append(" files, ")
        .append(committedSeedStats.totalBytes())
        .append(" bytes")
        .append(System.lineSeparator());
    builder.append("Newest Local Corpus Entries").append(System.lineSeparator());
    if (newestEntries.isEmpty()) {
      builder.append("  (none)").append(System.lineSeparator());
    } else {
      newestEntries.forEach(
          path -> builder.append("  - ").append(path).append(System.lineSeparator()));
    }
    builder.append("Committed Custom Seeds").append(System.lineSeparator());
    if (promotedInputs.isEmpty()) {
      builder.append("  (none)").append(System.lineSeparator());
    } else {
      promotedInputs.forEach(
          path -> builder.append("  - ").append(path).append(System.lineSeparator()));
    }
    if (!orphanedInputs.isEmpty()) {
      builder
          .append("WARNING: Seeds Without Promotion Metadata (")
          .append(orphanedInputs.size())
          .append(")")
          .append(System.lineSeparator());
      orphanedInputs.forEach(
          path -> builder.append("  ! ").append(path).append(System.lineSeparator()));
      builder
          .append("  Run 'jazzer/bin/promote <target> <input-path> <name>' for each listed seed.")
          .append(System.lineSeparator());
    }
    return builder.toString().trim();
  }

  /** Renders the currently recorded replayed-finding artifacts for one target. */
  static String renderFindingListing(String targetKey, List<FindingArtifact> findings) {
    StringBuilder builder = new StringBuilder();
    builder.append("Finding Listing").append(System.lineSeparator());
    builder.append("Target: ").append(targetKey).append(System.lineSeparator());
    if (findings.isEmpty()) {
      builder.append("No findings recorded.");
      return builder.toString();
    }
    findings.forEach(
        finding ->
            builder
                .append("- ")
                .append(finding.rawArtifactName())
                .append(" [")
                .append(finding.replayOutcome())
                .append("]")
                .append(System.lineSeparator())
                .append("  raw: ")
                .append(finding.rawArtifactPath())
                .append(System.lineSeparator())
                .append("  json: ")
                .append(finding.jsonArtifactPath())
                .append(System.lineSeparator())
                .append("  text: ")
                .append(finding.textArtifactPath())
                .append(System.lineSeparator()));
    return builder.toString().trim();
  }

  private static String renderMetrics(RunMetrics metrics) {
    return switch (metrics) {
      case RunMetrics.ActiveFuzzMetrics active ->
          "Executions: "
              + active.executions()
              + System.lineSeparator()
              + "Coverage: "
              + active.coverage()
              + System.lineSeparator()
              + "Features: "
              + active.features()
              + System.lineSeparator()
              + "Exec/s: "
              + active.executionsPerSecond()
              + System.lineSeparator()
              + "RSS: "
              + active.rssMegabytes()
              + " MB"
              + System.lineSeparator()
              + "Corpus Entries: "
              + active.corpusEntries()
              + System.lineSeparator()
              + "Corpus Bytes: "
              + active.corpusBytes()
              + System.lineSeparator()
              + "Max Input Bytes: "
              + active.maxInputBytes();
      case RunMetrics.RegressionMetrics regression ->
          "Executed Harnesses: " + regression.executedHarnessCount();
    };
  }

  private static String renderHarnessTelemetry(HarnessTelemetrySnapshot snapshot) {
    StringBuilder builder = new StringBuilder();
    builder
        .append("- ")
        .append(snapshot.displayName())
        .append(" [")
        .append(snapshot.harnessKey())
        .append("]")
        .append(System.lineSeparator());
    builder
        .append("  iterations=")
        .append(snapshot.iterations())
        .append(", success=")
        .append(snapshot.successfulOutcomes())
        .append(", expectedInvalid=")
        .append(snapshot.expectedInvalidOutcomes())
        .append(", generatedInvalid=")
        .append(snapshot.generatedInvalidInputs())
        .append(", unexpectedFailure=")
        .append(snapshot.unexpectedFailures())
        .append(System.lineSeparator());
    appendTopMap(builder, "  sequenceKinds", snapshot.sequenceKinds());
    appendTopMap(builder, "  readKinds", snapshot.readKinds());
    appendTopMap(builder, "  styleKinds", snapshot.styleKinds());
    appendTopMap(builder, "  sourceKinds", snapshot.sourceKinds());
    appendTopMap(builder, "  persistenceKinds", snapshot.persistenceKinds());
    appendTopMap(builder, "  responseKinds", snapshot.responseKinds());
    appendTopMap(builder, "  errorKinds", snapshot.errorKinds());
    return builder.toString();
  }

  private static String renderReplayDetails(ReplayDetails details) {
    return switch (details) {
      case ProtocolRequestDetails request ->
          "Input Bytes: "
              + request.inputBytes()
              + System.lineSeparator()
              + "Decode Outcome: "
              + request.decodeOutcome()
              + System.lineSeparator()
              + "Source Kind: "
              + request.sourceKind()
              + System.lineSeparator()
              + "Persistence Kind: "
              + request.persistenceKind()
              + System.lineSeparator()
              + "Operation Count: "
              + request.operationCount()
              + System.lineSeparator()
              + "Assertion Count: "
              + request.assertionCount()
              + System.lineSeparator()
              + "Read Count: "
              + request.readCount()
              + System.lineSeparator()
              + "Assertion Kinds: "
              + renderMapInline(request.assertionKinds())
              + System.lineSeparator()
              + "Read Kinds: "
              + renderMapInline(request.readKinds())
              + System.lineSeparator()
              + "Style Kinds: "
              + renderMapInline(request.styleKinds())
              + System.lineSeparator()
              + "Operation Kinds: "
              + renderMapInline(request.operationKinds());
      case ProtocolWorkflowDetails workflow ->
          "Input Bytes: "
              + workflow.inputBytes()
              + System.lineSeparator()
              + "Source Kind: "
              + workflow.sourceKind()
              + System.lineSeparator()
              + "Persistence Kind: "
              + workflow.persistenceKind()
              + System.lineSeparator()
              + "Operation Count: "
              + workflow.operationCount()
              + System.lineSeparator()
              + "Assertion Count: "
              + workflow.assertionCount()
              + System.lineSeparator()
              + "Read Count: "
              + workflow.readCount()
              + System.lineSeparator()
              + "Assertion Kinds: "
              + renderMapInline(workflow.assertionKinds())
              + System.lineSeparator()
              + "Read Kinds: "
              + renderMapInline(workflow.readKinds())
              + System.lineSeparator()
              + "Response Kind: "
              + workflow.responseKind()
              + System.lineSeparator()
              + "Style Kinds: "
              + renderMapInline(workflow.styleKinds())
              + System.lineSeparator()
              + "Operation Kinds: "
              + renderMapInline(workflow.operationKinds());
      case CommandSequenceDetails commandSequence ->
          "Input Bytes: "
              + commandSequence.inputBytes()
              + System.lineSeparator()
              + "Command Count: "
              + commandSequence.commandCount()
              + System.lineSeparator()
              + "Style Kinds: "
              + renderMapInline(commandSequence.styleKinds())
              + System.lineSeparator()
              + "Command Kinds: "
              + renderMapInline(commandSequence.commandKinds());
      case XlsxRoundTripDetails roundTrip ->
          "Input Bytes: "
              + roundTrip.inputBytes()
              + System.lineSeparator()
              + "Command Count: "
              + roundTrip.commandCount()
              + System.lineSeparator()
              + "Style Kinds: "
              + renderMapInline(roundTrip.styleKinds())
              + System.lineSeparator()
              + "Command Kinds: "
              + renderMapInline(roundTrip.commandKinds());
    };
  }

  private static void appendTopMap(StringBuilder builder, String label, Map<String, Long> values) {
    if (values == null || values.isEmpty()) {
      return;
    }
    builder
        .append(label)
        .append("=")
        .append(renderMapInline(values))
        .append(System.lineSeparator());
  }

  private static String renderMapInline(Map<String, Long> values) {
    if (values == null || values.isEmpty()) {
      return "(none)";
    }
    return values.entrySet().stream()
        .sorted(Map.Entry.comparingByValue(Comparator.reverseOrder()))
        .map(entry -> entry.getKey() + "=" + entry.getValue())
        .reduce((left, right) -> left + ", " + right)
        .orElse("(none)");
  }

  private static long actionableFindingCount(List<FindingArtifact> findings) {
    return findings.stream()
        .filter(finding -> finding.replayOutcome().equals("UNEXPECTED_FAILURE"))
        .count();
  }

  private static long expectedInvalidArtifactCount(List<FindingArtifact> findings) {
    return findings.stream()
        .filter(finding -> finding.replayOutcome().equals("EXPECTED_INVALID"))
        .count();
  }

  private static long replayCleanArtifactCount(List<FindingArtifact> findings) {
    return findings.stream().filter(finding -> finding.replayOutcome().equals("SUCCESS")).count();
  }
}
