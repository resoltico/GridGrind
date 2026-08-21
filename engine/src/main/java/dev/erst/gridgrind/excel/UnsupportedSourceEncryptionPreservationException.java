package dev.erst.gridgrind.excel;

/** Signals that a readable source encryption envelope cannot be reapplied on the write path. */
public final class UnsupportedSourceEncryptionPreservationException
    extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  /** Creates one precise unsupported-source-encryption failure. */
  public UnsupportedSourceEncryptionPreservationException(String message) {
    super(message);
  }
}
