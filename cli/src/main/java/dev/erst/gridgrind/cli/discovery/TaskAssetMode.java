package dev.erst.gridgrind.cli.discovery;

/** Typed asset dependency posture for one CLI-owned task descriptor. */
public enum TaskAssetMode {
  SELF_CONTAINED,
  REQUIRES_EXTERNAL_PAYLOADS
}
