package dev.erst.gridgrind.cli;

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
  public static void main(String[] args) throws IOException {
    new App().run(args, System.in, System.out, System.err);
  }

  void run(String[] args) throws IOException {
    run(args, System.in, System.out, System.err);
  }

  void run(String[] args, InputStream stdin, OutputStream stdout, OutputStream stderr)
      throws IOException {
    int exitCode =
        cliFactory
            .create()
            .run(
                args,
                Objects.requireNonNull(stdin, "stdin must not be null"),
                Objects.requireNonNull(stdout, "stdout must not be null"),
                Objects.requireNonNull(stderr, "stderr must not be null"));
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
