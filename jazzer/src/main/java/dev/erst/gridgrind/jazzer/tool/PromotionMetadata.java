package dev.erst.gridgrind.jazzer.tool;

/** Describes one committed regression input promoted from a local Jazzer artifact or corpus file. */
public record PromotionMetadata(
    String targetKey,
    String sourcePath,
    String promotedInputPath,
    String replayOutcome,
    String promotedAt,
    String replayTextPath) {}
