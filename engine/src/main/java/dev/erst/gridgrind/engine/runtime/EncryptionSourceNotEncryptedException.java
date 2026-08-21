package dev.erst.gridgrind.engine.runtime;

/** Signals that declared source-encryption preservation has no encrypted source to preserve. */
final class EncryptionSourceNotEncryptedException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  EncryptionSourceNotEncryptedException() {
    super("PRESERVE_SOURCE encryption requires an encrypted source workbook");
  }
}
