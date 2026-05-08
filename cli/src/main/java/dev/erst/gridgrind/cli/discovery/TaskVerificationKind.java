package dev.erst.gridgrind.cli.discovery;

/** Typed verification postures advertised by one CLI-owned task descriptor. */
public enum TaskVerificationKind {
  FACT_READBACK,
  ASSERTION_CHECKS,
  HEALTH_ANALYSIS,
  EXPORT_REREAD
}
