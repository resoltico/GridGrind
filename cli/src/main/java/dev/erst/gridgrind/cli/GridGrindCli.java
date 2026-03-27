package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.protocol.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.protocol.GridGrindProblemCode;
import dev.erst.gridgrind.protocol.GridGrindProblems;
import dev.erst.gridgrind.protocol.GridGrindProtocolVersion;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindRequestExecutor;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.InvalidJsonException;
import dev.erst.gridgrind.protocol.InvalidRequestException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Objects;

/** Thin command-line transport for the GridGrind protocol. */
public final class GridGrindCli {
  private final GridGrindRequestExecutor requestExecutor;
  private final CliRequestReader requestReader;
  private final CliResponseWriter responseWriter;

  /** Creates the production CLI backed by the default request executor and transport helpers. */
  public GridGrindCli() {
    this(new DefaultGridGrindRequestExecutor(), new CliRequestReader(), new CliResponseWriter());
  }

  GridGrindCli(GridGrindRequestExecutor requestExecutor) {
    this(requestExecutor, new CliRequestReader(), new CliResponseWriter());
  }

  GridGrindCli(
      GridGrindRequestExecutor requestExecutor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter) {
    this.requestExecutor = GridGrindRequestExecutor.requireNonNull(requestExecutor);
    this.requestReader = Objects.requireNonNull(requestReader, "requestReader must not be null");
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
  }

  /** Runs one CLI invocation against stdin/stdout or explicit request/response file paths. */
  public int run(String[] args, InputStream stdin, OutputStream stdout) throws IOException {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");

    CliCommand command;
    try {
      command = CliArguments.parse(args);
    } catch (CliArgumentsException exception) {
      responseWriter.write(
          stdout,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              exception.getMessage(),
              new GridGrindResponse.ProblemContext.ParseArguments(exception.argument()),
              exception));
      return 2;
    } catch (IllegalArgumentException exception) {
      responseWriter.write(
          stdout,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message(exception),
              new GridGrindResponse.ProblemContext.ParseArguments(null),
              exception));
      return 2;
    }

    return switch (command) {
      case CliCommand.Version _ -> {
        stdout.write(("gridgrind " + version() + "\n").getBytes(StandardCharsets.UTF_8));
        stdout.flush();
        yield 0;
      }
      case CliCommand.Execute execute -> executeCommand(execute, stdin, stdout);
    };
  }

  private int executeCommand(CliCommand.Execute command, InputStream stdin, OutputStream stdout)
      throws IOException {
    GridGrindRequest request;
    try {
      request = requestReader.read(command.requestPath(), stdin);
    } catch (InvalidRequestException exception) {
      return responseWriter.write(
          command.responsePath(),
          stdout,
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(command.requestPath()), null, null, null))));
    } catch (InvalidJsonException exception) {
      return responseWriter.write(
          command.responsePath(),
          stdout,
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(command.requestPath()), null, null, null))));
    } catch (IOException exception) {
      return responseWriter.write(
          command.responsePath(),
          stdout,
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(command.requestPath()), null, null, null))));
    }

    GridGrindResponse response;
    try {
      response = requestExecutor.execute(request);
    } catch (Exception exception) {
      response =
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception, new GridGrindResponse.ProblemContext.ExecuteRequest(null, null)));
    }

    return responseWriter.write(command.responsePath(), stdout, response);
  }

  private static String version() {
    return versionFrom(GridGrindCli.class.getPackage().getImplementationVersion());
  }

  /**
   * Returns the given implementation version string, or {@code "unknown"} when the JAR manifest
   * attribute is absent (e.g. when running from the test classpath without a packaged JAR).
   */
  static String versionFrom(String implementationVersion) {
    if (implementationVersion == null) {
      return "unknown";
    }
    return implementationVersion;
  }

  private GridGrindResponse.Failure failure(
      GridGrindProblemCode code,
      String message,
      GridGrindResponse.ProblemContext context,
      Throwable cause) {
    return new GridGrindResponse.Failure(
        GridGrindProtocolVersion.current(),
        GridGrindProblems.problem(code, message, context, cause));
  }

  private String message(Exception exception) {
    return exception.getMessage();
  }

  private String pathString(Path path) {
    return path == null ? null : path.toAbsolutePath().toString();
  }
}
