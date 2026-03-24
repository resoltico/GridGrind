package dev.erst.gridgrind.cli;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/** Process entry point for the GridGrind CLI. */
public final class App {
  static CliFactory cliFactory = () -> new AgentCli()::run;
  static ExitHandler exitHandler = System::exit;

  private App() {}

  /** Entry point: creates the CLI runner and exits with a non-zero code on failure. */
  public static void main(String[] args) throws IOException {
    int exitCode = cliFactory.create().run(args, System.in, System.out);
    if (exitCode != 0) {
      exitHandler.exit(exitCode);
    }
  }

  /** Functional interface for running one CLI invocation. */
  @FunctionalInterface
  interface CliRunner {
    int run(String[] args, InputStream stdin, OutputStream stdout) throws IOException;
  }

  /** Functional interface for creating a CliRunner. */
  @FunctionalInterface
  interface CliFactory {
    CliRunner create();
  }

  /** Functional interface for terminating the process with an exit code. */
  @FunctionalInterface
  interface ExitHandler {
    void exit(int exitCode);
  }
}
