package dev.erst.gridgrind.core.result;

/** Structured execution-journal detail levels for response telemetry and host rendering. */
public enum ExecutionJournalLevel {
  /** Compact target summaries and no live progress stream. */
  SUMMARY,
  /** Expanded target summaries and no live progress stream. */
  NORMAL,
  /** Expanded target summaries plus live structured progress on the execution sink. */
  VERBOSE
}
