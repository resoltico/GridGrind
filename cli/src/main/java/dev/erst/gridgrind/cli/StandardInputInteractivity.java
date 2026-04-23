package dev.erst.gridgrind.cli;

import java.io.Console;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BooleanSupplier;

/**
 * Detects whether the current process stdin is attached to an interactive terminal.
 *
 * <p>{@link System#console()} covers the common fully-interactive case but returns {@code null}
 * when stdout is redirected. A separate native stdin probe keeps no-argument help working for
 * invocations such as {@code java -jar gridgrind.jar > help.txt}.
 */
final class StandardInputInteractivity implements BooleanSupplier {
  private static final StandardInputInteractivity NEVER =
      new StandardInputInteractivity(() -> false, NativeStandardInputProbe.unsupported());

  private final BooleanSupplier consoleIsInteractive;
  private final BooleanSupplier nativeStandardInputProbe;

  static StandardInputInteractivity currentProcess() {
    return new StandardInputInteractivity(
        StandardInputInteractivity::consoleIsInteractive,
        NativeStandardInputProbe.currentProcess());
  }

  static StandardInputInteractivity never() {
    return NEVER;
  }

  @SuppressWarnings("SystemConsoleNull")
  static boolean consoleIsInteractive() {
    return Optional.ofNullable(System.console()).map(Console::isTerminal).orElse(false);
  }

  static boolean consoleIsInteractive(
      boolean consolePresent, BooleanSupplier terminalIsInteractive) {
    return consolePresent && terminalIsInteractive.getAsBoolean();
  }

  StandardInputInteractivity(
      BooleanSupplier consoleIsInteractive, BooleanSupplier nativeStandardInputProbe) {
    this.consoleIsInteractive =
        Objects.requireNonNull(consoleIsInteractive, "consoleIsInteractive must not be null");
    this.nativeStandardInputProbe =
        Objects.requireNonNull(
            nativeStandardInputProbe, "nativeStandardInputProbe must not be null");
  }

  @Override
  public boolean getAsBoolean() {
    return consoleIsInteractive.getAsBoolean() || nativeStandardInputProbe.getAsBoolean();
  }
}
