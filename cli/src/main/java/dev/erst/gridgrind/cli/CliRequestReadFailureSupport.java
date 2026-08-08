package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.json.RequestDiagnosticRedactor;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Builds and routes rejected command diagnostics for failures before workbook execution begins. */
final class CliRequestReadFailureSupport {
  private CliRequestReadFailureSupport() {}

  static RequestInput requestInput(Optional<Path> requestPath) {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    return requestPath.isEmpty() || CliPathArguments.isStandardInputPath(requestPath)
        ? RequestInput.standardInput()
        : RequestInput.requestFile(requestPath.orElseThrow().toAbsolutePath().toString());
  }

  static int write(
      CliResponseWriter responseWriter,
      String command,
      Optional<Path> requestPath,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      Throwable exception,
      Optional<RequestDiagnosticRedactor> requestRedactor,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(responseWriter, "responseWriter must not be null");
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(requestRedactor, "requestRedactor must not be null");
    var diagnostic =
        CommandErrors.readRequestFailure(
            command,
            GridGrindProblems.fromException(
                exception,
                new ProblemContext.ReadRequest(
                    requestInput(requestPath),
                    CliRequestFailureLocationSupport.locationFor(exception))));
    if (requestRedactor.isPresent()) {
      return responseWriter.writeCommandError(
          responsePath, stdout, stderr, diagnostic, requestRedactor.orElseThrow(), prettyJson);
    }
    return responseWriter.writeCommandError(responsePath, stdout, stderr, diagnostic, prettyJson);
  }
}
