package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.cli.discovery.CliTransport;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Last-resort structured failure emission for unexpected CLI/runtime faults. */
final class CliUnexpectedFailureSupport {
  private CliUnexpectedFailureSupport() {}

  static int emit(
      String[] args,
      Optional<Path> responsePathHint,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      Throwable exception) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(responsePathHint, "responsePathHint must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(exception, "exception must not be null");

    CliDiagnostic diagnostic =
        CliDiagnostics.unexpectedFailure(
            CliPrimaryCommandSupport.primaryCommandName(args), exception);
    try {
      return new CliResponseWriter()
          .writeCliDiagnostic(responsePathHint, stdout, stderr, diagnostic, prettyJson);
    } catch (Throwable writeFailure) {
      directFallback(
          diagnostic, stdout, stderr, responsePathHint.isPresent(), prettyJson, writeFailure);
      return diagnostic.exitCode();
    }
  }

  private static void directFallback(
      CliDiagnostic diagnostic,
      OutputStream stdout,
      OutputStream stderr,
      boolean responsePathUsed,
      boolean prettyJson,
      Throwable writeFailure) {
    CliDiagnostic stdoutDiagnostic =
        responsePathUsed
            ? CliResponseWriter.diagnosticWithTransport(diagnostic, CliTransport.standardOutput())
            : diagnostic;
    try {
      byte[] payload = GridGrindCliJson.writeBytes(stdoutDiagnostic, prettyJson);
      writePayload(stdout, payload);
      return;
    } catch (IOException exception) {
      if (structuredStderrWasAlreadyRecovered(responsePathUsed, writeFailure)) {
        return;
      }
      try {
        writePayload(stderr, GridGrindCliJson.writeBytes(diagnostic, prettyJson));
        return;
      } catch (IOException exceptionOnStderr) {
        exception.addSuppressed(exceptionOnStderr);
      }
      try {
        String message =
            responsePathUsed
                ? "GridGrind failed before it could emit a structured error payload to the response fallback channels."
                : "GridGrind failed before it could emit a structured error payload.";
        stderr.write(
            (message + System.lineSeparator()).getBytes(java.nio.charset.StandardCharsets.UTF_8));
        stderr.flush();
      } catch (IOException humanReadableFailure) {
        exception.addSuppressed(humanReadableFailure);
      }
    }
  }

  private static boolean structuredStderrWasAlreadyRecovered(
      boolean responsePathUsed, Throwable writeFailure) {
    if (!responsePathUsed || !(writeFailure instanceof IOException ioFailure)) {
      return false;
    }
    return ioFailure.getSuppressed().length == 0;
  }

  private static void writePayload(OutputStream outputStream, byte[] payload) throws IOException {
    CliPayloadOutput.write(outputStream, payload);
  }
}
