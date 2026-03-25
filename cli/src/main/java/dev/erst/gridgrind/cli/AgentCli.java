package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.protocol.GridGrindJson;
import dev.erst.gridgrind.protocol.GridGrindProblemCode;
import dev.erst.gridgrind.protocol.GridGrindProblems;
import dev.erst.gridgrind.protocol.GridGrindProtocolVersion;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import dev.erst.gridgrind.protocol.GridGrindResponse;
import dev.erst.gridgrind.protocol.GridGrindService;
import dev.erst.gridgrind.protocol.InvalidJsonException;
import dev.erst.gridgrind.protocol.InvalidRequestException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** Thin command-line transport for the GridGrind protocol. */
public final class AgentCli {
  private final GridGrindHandler gridGrindHandler;

  /** Creates the production AgentCli backed by the default {@link GridGrindService}. */
  public AgentCli() {
    this(new GridGrindService()::execute);
  }

  AgentCli(GridGrindHandler gridGrindHandler) {
    this.gridGrindHandler =
        Objects.requireNonNull(gridGrindHandler, "gridGrindHandler must not be null");
  }

  /** Runs one CLI invocation against stdin/stdout or explicit request/response file paths. */
  public int run(String[] args, InputStream stdin, OutputStream stdout) throws IOException {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");

    CliOptions options;
    try {
      options = CliOptions.parse(args);
    } catch (CliArgumentException exception) {
      writeResponse(
          stdout,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              exception.getMessage(),
              new GridGrindResponse.ProblemContext.ParseArguments(exception.argument()),
              exception));
      return 2;
    } catch (IllegalArgumentException exception) {
      writeResponse(
          stdout,
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message(exception),
              new GridGrindResponse.ProblemContext.ParseArguments(null),
              exception));
      return 2;
    }

    GridGrindRequest request;
    try {
      request = readRequest(options.requestPath(), stdin);
    } catch (InvalidRequestException exception) {
      return writeResponseWithFallback(
          options.responsePath(),
          stdout,
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(options.requestPath()), null, null, null))));
    } catch (InvalidJsonException exception) {
      return writeResponseWithFallback(
          options.responsePath(),
          stdout,
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(options.requestPath()), null, null, null))));
    } catch (IOException exception) {
      return writeResponseWithFallback(
          options.responsePath(),
          stdout,
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new GridGrindResponse.ProblemContext.ReadRequest(
                      pathString(options.requestPath()), null, null, null))));
    }

    GridGrindResponse response;
    try {
      response = gridGrindHandler.execute(request);
    } catch (Exception exception) {
      response =
          new GridGrindResponse.Failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception, new GridGrindResponse.ProblemContext.ExecuteRequest(null, null)));
    }

    return writeResponseWithFallback(options.responsePath(), stdout, response);
  }

  private GridGrindRequest readRequest(Path requestPath, InputStream stdin) throws IOException {
    if (requestPath == null) {
      return GridGrindJson.readRequest(stdin);
    }
    try (InputStream requestInput = Files.newInputStream(requestPath)) {
      return GridGrindJson.readRequest(requestInput);
    }
  }

  private int writeResponseWithFallback(
      Path responsePath, OutputStream stdout, GridGrindResponse response) throws IOException {
    if (responsePath == null) {
      writeResponse(stdout, response);
      return exitCodeFor(response);
    }

    try {
      Path parent = responsePath.getParent();
      if (parent != null) {
        Files.createDirectories(parent);
      }
      try (OutputStream responseOutput = Files.newOutputStream(responsePath)) {
        writeResponse(responseOutput, response);
      }
      return exitCodeFor(response);
    } catch (IOException exception) {
      GridGrindResponse.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new GridGrindResponse.ProblemContext.WriteResponse(pathString(responsePath)));
      if (response instanceof GridGrindResponse.Failure f) {
        problem =
            GridGrindProblems.appendCause(problem, GridGrindProblems.problemCause(f.problem()));
      }
      writeResponse(
          stdout, new GridGrindResponse.Failure(GridGrindProtocolVersion.current(), problem));
      return 1;
    }
  }

  private void writeResponse(OutputStream outputStream, GridGrindResponse response)
      throws IOException {
    GridGrindJson.writeResponse(outputStream, response);
    outputStream.write('\n');
    outputStream.flush();
  }

  private int exitCodeFor(GridGrindResponse response) {
    return switch (response) {
      case GridGrindResponse.Success _ -> 0;
      case GridGrindResponse.Failure _ -> 1;
    };
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

  private record CliOptions(Path requestPath, Path responsePath) {
    static CliOptions parse(String[] args) {
      Path requestPath = null;
      Path responsePath = null;

      int index = 0;
      while (index < args.length) {
        String argument = args[index];
        switch (argument) {
          case "--request" -> {
            if (requestPath != null) {
              throw new CliArgumentException("--request", "Duplicate argument: --request");
            }
            int valueIndex = nextValueIndex(args, index, "--request");
            requestPath = Path.of(args[valueIndex]);
            index = valueIndex + 1;
          }
          case "--response" -> {
            if (responsePath != null) {
              throw new CliArgumentException("--response", "Duplicate argument: --response");
            }
            int valueIndex = nextValueIndex(args, index, "--response");
            responsePath = Path.of(args[valueIndex]);
            index = valueIndex + 1;
          }
          default -> throw new CliArgumentException(argument, "Unknown argument: " + argument);
        }
      }

      return new CliOptions(requestPath, responsePath);
    }

    private static int nextValueIndex(String[] args, int flagIndex, String flagName) {
      int valueIndex = flagIndex + 1;
      if (valueIndex >= args.length) {
        throw new CliArgumentException(flagName, "Missing value for " + flagName);
      }
      return valueIndex;
    }
  }

  /** Signals an unrecognized or malformed CLI argument. */
  private static final class CliArgumentException extends IllegalArgumentException {
    private static final long serialVersionUID = 1L;

    private final String argument;

    private CliArgumentException(String argument, String message) {
      super(message);
      this.argument = argument;
    }

    private String argument() {
      return argument;
    }
  }

  /** Functional interface for executing a {@link GridGrindRequest} and returning a response. */
  @FunctionalInterface
  interface GridGrindHandler {
    /** Executes the request and returns the corresponding response. */
    GridGrindResponse execute(GridGrindRequest request);
  }
}
