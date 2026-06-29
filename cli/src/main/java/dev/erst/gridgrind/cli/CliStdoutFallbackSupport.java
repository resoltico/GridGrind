package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Objects;

/** Shared stdout-fallback payload rendering and stderr notice formatting. */
final class CliStdoutFallbackSupport {
  private CliStdoutFallbackSupport() {}

  static StdoutFallback cliFailureReport(String description, CliFailureReport report)
      throws IOException {
    return new StdoutFallback(description, GridGrindCliJson.writeBytes(report));
  }

  static StdoutFallback response(String description, GridGrindResponse response)
      throws IOException {
    return new StdoutFallback(description, GridGrindJson.writeResponseBytes(response));
  }

  static StdoutFallback doctorReport(String description, RequestDoctorReport report)
      throws IOException {
    return new StdoutFallback(description, GridGrindJson.writeRequestDoctorReportBytes(report));
  }

  static void write(
      OutputStream stderr,
      OutputStream stdout,
      IOException exception,
      Path targetPath,
      StdoutFallback fallback)
      throws IOException {
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(fallback, "fallback must not be null");
    String line =
        CliResponseWriter.responseWriteMessage(exception, targetPath)
            + ". Wrote the "
            + fallback.description()
            + " to stdout instead."
            + System.lineSeparator();
    stderr.write(line.getBytes(StandardCharsets.UTF_8));
    stderr.flush();
    CliPayloadOutput.write(stdout, fallback.payload());
  }

  /** Value object for a stdout fallback payload and its operator-facing description. */
  static final class StdoutFallback {
    private final String description;
    private final byte[] payload;

    private StdoutFallback(String description, byte[] payload) {
      this.description = Objects.requireNonNull(description, "description must not be null");
      this.payload = Objects.requireNonNull(payload, "payload must not be null").clone();
    }

    String description() {
      return description;
    }

    byte[] payload() {
      return payload.clone();
    }
  }
}
