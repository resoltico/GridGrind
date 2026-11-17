package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.command.CliArguments;
import dev.erst.gridgrind.cli.command.CliPathArguments;
import dev.erst.gridgrind.cli.command.CliUnexpectedFailureSupport;
import dev.erst.gridgrind.cli.command.GridGrindCli;
import dev.erst.gridgrind.cli.io.CliPrimaryOutputException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Objects;

/** Process entry point for the GridGrind CLI. */
public final class App {
  private final CliFactory cliFactory;
  private final ExitHandler exitHandler;

  /** Creates the production App wired to the default CLI factory and {@code System::exit}. */
  public App() {
    this(new ProductionCliFactory(), System::exit);
  }

  App(CliFactory cliFactory, ExitHandler exitHandler) {
    this.cliFactory = Objects.requireNonNull(cliFactory, "cliFactory must not be null");
    this.exitHandler = Objects.requireNonNull(exitHandler, "exitHandler must not be null");
  }

  /** Entry point: creates an App with production defaults and runs the CLI. */
  @SuppressWarnings({"PMD.UseProperClassLoader", "PMD.UnusedLocalVariable"})
  // The shipped provider belongs to App, never an embedding TCCL; logging is a close-only owner.
  public static void main(String[] args) {
    var diagnostics = new dev.erst.gridgrind.cli.io.CliDiagnosticChannel(System.err);
    String threshold;
    try {
      threshold =
          dev.erst.gridgrind.cli.logging.StandaloneLoggingPreflight.validate(
              App.class.getClassLoader());
    } catch (dev.erst.gridgrind.cli.logging.StandaloneLoggingRefusal refusal) {
      int code =
          dev.erst.gridgrind.cli.command.CliStandaloneFailureSupport.emit(
              args, System.out, diagnostics, refusal);
      if (code != 0) System.exit(code);
      return;
    } catch (Throwable failure) {
      int code =
          CliUnexpectedFailureSupport.emit(
              args,
              CliPathArguments.responsePathHint(args),
              CliArguments.prettyJsonHint(args),
              System.out,
              diagnostics,
              failure);
      if (code != 0) System.exit(code);
      return;
    }
    var result = new java.util.concurrent.atomic.AtomicInteger();
    boolean invocationStarted = false;
    try (var logging =
        dev.erst.gridgrind.cli.logging.StandaloneLogging.open(
            App.class.getClassLoader(), threshold, diagnostics)) {
      invocationStarted = true;
      new App(new ProductionCliFactory(), result::set)
          .run(args, System.in, System.out, diagnostics);
    } catch (Throwable failure) {
      if (invocationStarted) result.set(1);
      else
        result.set(
            CliUnexpectedFailureSupport.emit(
                args,
                CliPathArguments.responsePathHint(args),
                CliArguments.prettyJsonHint(args),
                System.out,
                diagnostics,
                failure));
    }
    if (result.get() != 0) System.exit(result.get());
  }

  void run(String... args) {
    run(args, System.in, System.out, System.err);
  }

  void run(String[] args, InputStream stdin, OutputStream stdout, OutputStream stderr) {
    int exitCode;
    try {
      exitCode =
          cliFactory
              .create()
              .run(
                  args,
                  Objects.requireNonNull(stdin, "stdin must not be null"),
                  Objects.requireNonNull(stdout, "stdout must not be null"),
                  Objects.requireNonNull(stderr, "stderr must not be null"));
    } catch (CliPrimaryOutputException ignored) {
      exitCode = 1;
    } catch (Throwable exception) {
      exitCode =
          CliUnexpectedFailureSupport.emit(
              args,
              CliPathArguments.responsePathHint(args),
              CliArguments.prettyJsonHint(args),
              Objects.requireNonNull(stdout, "stdout must not be null"),
              Objects.requireNonNull(stderr, "stderr must not be null"),
              exception);
    }
    if (exitCode != 0) {
      exitHandler.exit(exitCode);
    }
  }

  /** Functional interface for running one CLI invocation. */
  @FunctionalInterface
  interface CliRunner {
    /** Runs the CLI with the given args and streams, returning an exit code. */
    int run(String[] args, InputStream stdin, OutputStream stdout, OutputStream stderr)
        throws IOException;
  }

  /** Functional interface for creating a {@link CliRunner}. */
  @FunctionalInterface
  interface CliFactory {
    /** Creates a {@link CliRunner} for one invocation. */
    CliRunner create();
  }

  /** Functional interface for terminating the process with an exit code. */
  @FunctionalInterface
  interface ExitHandler {
    /** Terminates the process with the given exit code. */
    void exit(int exitCode);
  }

  /** Creates the production CLI runner from the repository-owned CLI entry point. */
  private static final class ProductionCliFactory implements CliFactory {
    @Override
    public CliRunner create() {
      GridGrindCli cli = new GridGrindCli();
      return cli::run;
    }
  }
}
