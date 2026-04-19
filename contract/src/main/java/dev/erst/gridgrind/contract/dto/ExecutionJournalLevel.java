package dev.erst.gridgrind.contract.dto;

/** Structured execution-journal detail levels for response telemetry and CLI rendering. */
public enum ExecutionJournalLevel {
  /** Compact target summaries and no live event stream. */
  SUMMARY,
  /** Expanded target summaries and no live event stream. */
  NORMAL,
  /** Expanded target summaries plus fine-grained execution events. */
  VERBOSE
}
