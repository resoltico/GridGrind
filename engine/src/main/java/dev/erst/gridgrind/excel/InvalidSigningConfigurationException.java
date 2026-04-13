package dev.erst.gridgrind.excel;

/** Signals that the supplied OOXML signing configuration cannot load a signing key and chain. */
public final class InvalidSigningConfigurationException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  /** Creates the exception for the given signing-configuration problem. */
  public InvalidSigningConfigurationException(String message) {
    super(message);
  }

  /** Creates the exception for the given signing-configuration problem with a cause. */
  public InvalidSigningConfigurationException(String message, Throwable cause) {
    super(message, cause);
  }
}
