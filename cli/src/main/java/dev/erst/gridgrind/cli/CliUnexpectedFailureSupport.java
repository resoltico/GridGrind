package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.CommandError;
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

    CommandError diagnostic =
        CommandErrors.unexpectedFailure(
            CliPrimaryCommandSupport.primaryCommandName(args), exception);
    try {
      return new CliResponseWriter()
          .writeCommandError(responsePathHint, stdout, stderr, diagnostic, prettyJson);
    } catch (Throwable ignored) {
      directFallback(diagnostic, stdout, stderr, responsePathHint, prettyJson);
      return CliExitCodes.forCommandError(diagnostic);
    }
  }

  static void directFallback(
      CommandError diagnostic,
      OutputStream stdout,
      OutputStream stderr,
      Optional<Path> responsePathHint,
      boolean prettyJson) {
    try {
      byte[] payload = GridGrindCliJson.writeBytes(diagnostic, prettyJson);
      writePayload(stdout, payload);
      if (responsePathHint.isPresent()) {
        try {
          CliResponseTransportSupport.writeTransportNoticeToStderr(
              stderr,
              CliTransportNotice.stdoutFallback(
                  responsePathHint.orElseThrow().toAbsolutePath().toString()));
        } catch (IOException ignored) {
          // The recovered command error on stdout remains the primary fallback result.
        }
      }
    } catch (IOException ignored) {
      // The primary payload is stdout-only; stderr must never carry a competing result schema.
    }
  }

  private static void writePayload(OutputStream outputStream, byte[] payload) throws IOException {
    CliPayloadOutput.write(outputStream, payload);
  }
}
