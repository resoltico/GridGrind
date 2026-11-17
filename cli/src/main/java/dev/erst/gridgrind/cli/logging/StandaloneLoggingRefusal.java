package dev.erst.gridgrind.cli.logging;

/** A known standalone logging setting or metadata role conflicts with the owned transport. */
public final class StandaloneLoggingRefusal extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;
  private final String setting;

  StandaloneLoggingRefusal(String setting) {
    super("Unsupported standalone logging configuration: " + setting);
    this.setting = setting;
  }

  StandaloneLoggingRefusal(String setting, Throwable cause) {
    super("Unsupported standalone logging configuration: " + setting, cause);
    this.setting = setting;
  }

  /** Safe setting identity or fixed packaged metadata role; never its configured value. */
  public String setting() {
    return setting;
  }
}
