package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
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

    CliFailureReport report =
        CliFailureReports.unexpectedFailure(
            CliPrimaryCommandSupport.primaryCommandName(args), "unexpected-failure", exception);
    try {
      return new CliResponseWriter()
          .writeCliFailureReport(responsePathHint, stdout, stderr, report, prettyJson);
    } catch (Throwable writeFailure) {
      directFallback(report, stdout, stderr, responsePathHint.isPresent(), prettyJson);
      return report.exitCode();
    }
  }

  private static void directFallback(
      CliFailureReport report,
      OutputStream stdout,
      OutputStream stderr,
      boolean responsePathUsed,
      boolean prettyJson) {
    try {
      byte[] payload = GridGrindCliJson.writeBytes(report, prettyJson);
      writePayload(stdout, payload);
      return;
    } catch (IOException exception) {
      try {
        writePayload(stderr, GridGrindCliJson.writeBytes(report, prettyJson));
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

  private static void writePayload(OutputStream outputStream, byte[] payload) throws IOException {
    CliPayloadOutput.write(outputStream, payload);
  }
}
