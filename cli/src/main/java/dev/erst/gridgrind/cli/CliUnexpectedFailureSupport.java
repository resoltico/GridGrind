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
          .writeCliFailureReport(responsePathHint, stdout, stderr, report);
    } catch (Throwable writeFailure) {
      directFallback(report, stdout, stderr, responsePathHint.isPresent());
      return report.exitCode();
    }
  }

  private static void directFallback(
      CliFailureReport report, OutputStream stdout, OutputStream stderr, boolean responsePathUsed) {
    try {
      byte[] payload = GridGrindCliJson.writeCliFailureReportBytes(report);
      if (responsePathUsed) {
        writePayload(stdout, payload);
      } else {
        writePayload(stderr, payload);
      }
    } catch (IOException ignored) {
      try {
        stderr.write(
            ("GridGrind failed before it could emit a structured error payload."
                    + System.lineSeparator())
                .getBytes(java.nio.charset.StandardCharsets.UTF_8));
        stderr.flush();
      } catch (IOException ignoredAgain) {
        return;
      }
    }
  }

  private static void writePayload(OutputStream outputStream, byte[] payload) throws IOException {
    CliPayloadOutput.write(outputStream, payload);
  }
}
