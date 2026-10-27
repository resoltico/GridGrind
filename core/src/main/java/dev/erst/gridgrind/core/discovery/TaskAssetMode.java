package dev.erst.gridgrind.core.discovery;

/** Typed asset dependency posture for one Product-owned task descriptor. */
public enum TaskAssetMode {
  SELF_CONTAINED,
  REQUIRES_EXTERNAL_PAYLOADS
}
