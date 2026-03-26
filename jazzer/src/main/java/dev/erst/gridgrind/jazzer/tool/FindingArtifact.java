package dev.erst.gridgrind.jazzer.tool;

/** Describes one replayed local finding artifact derived from a raw libFuzzer output file. */
public record FindingArtifact(
    String rawArtifactName,
    String rawArtifactPath,
    String replayOutcome,
    String jsonArtifactPath,
    String textArtifactPath) {}
