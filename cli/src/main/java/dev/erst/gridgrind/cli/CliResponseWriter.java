package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.executor.GridGrindProblems;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** Writes GridGrind responses to stdout or an explicit response file with structured fallback. */
final class CliResponseWriter {
  /** Writes the response to the configured destination and returns the corresponding exit code. */
  int write(Path responsePath, OutputStream stdout, GridGrindResponse response) throws IOException {
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(response, "response must not be null");
    if (responsePath == null) {
      write(stdout, response);
      return exitCodeFor(response);
    }

    try {
      Path parent = responsePath.getParent();
      if (parent != null) {
        Files.createDirectories(parent);
      }
      try (OutputStream responseOutput = Files.newOutputStream(responsePath)) {
        write(responseOutput, response);
      }
      return exitCodeFor(response);
    } catch (IOException exception) {
      GridGrindResponse.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new GridGrindResponse.ProblemContext.WriteResponse(
                  responsePath.toAbsolutePath().toString()));
      if (response instanceof GridGrindResponse.Failure failure) {
        problem =
            GridGrindProblems.appendCause(
                problem, GridGrindProblems.problemCause(failure.problem()));
      }
      write(stdout, new GridGrindResponse.Failure(GridGrindProtocolVersion.current(), problem));
      return 1;
    }
  }

  /** Writes one response to an already-open output stream, preserving caller stream ownership. */
  void write(OutputStream outputStream, GridGrindResponse response) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(response, "response must not be null");
    GridGrindJson.writeResponse(outputStream, response);
    outputStream.write('\n');
    outputStream.flush();
  }

  /** Returns the process exit code associated with the response shape. */
  static int exitCodeFor(GridGrindResponse response) {
    return switch (response) {
      case GridGrindResponse.Success _ -> 0;
      case GridGrindResponse.Failure _ -> 1;
    };
  }
}
