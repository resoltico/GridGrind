package dev.erst.gridgrind.jazzer.tool;

/** Captures the size of a local Jazzer corpus at one point in time. */
public record CorpusStats(long fileCount, long totalBytes) {}
