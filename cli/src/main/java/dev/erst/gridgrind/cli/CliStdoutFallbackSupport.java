package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Objects;

/** Shared stdout-fallback payload rendering and stderr notice formatting. */
final class CliStdoutFallbackSupport {
  private CliStdoutFallbackSupport() {}

  static StdoutFallback cliDiagnostic(CliDiagnostic diagnostic) throws IOException {
    return cliDiagnostic(diagnostic, false);
  }

  static StdoutFallback cliDiagnostic(CliDiagnostic diagnostic, boolean prettyJson)
      throws IOException {
    return new StdoutFallback(GridGrindCliJson.writeBytes(diagnostic, prettyJson));
  }

  static StdoutFallback response(GridGrindResponse response) throws IOException {
    return response(response, false);
  }

  static StdoutFallback response(GridGrindResponse response, boolean prettyJson)
      throws IOException {
    return new StdoutFallback(GridGrindJsonOutput.writeResponseBytes(response, prettyJson));
  }

  static StdoutFallback doctorReport(RequestDoctorReport report) throws IOException {
    return doctorReport(report, false);
  }

  static StdoutFallback doctorReport(RequestDoctorReport report, boolean prettyJson)
      throws IOException {
    return new StdoutFallback(
        GridGrindJsonOutput.writeRequestDoctorReportBytes(report, prettyJson));
  }

  static void write(
      OutputStream stderr,
      OutputStream stdout,
      CliDiagnostic stderrDiagnostic,
      StdoutFallback fallback,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderrDiagnostic, "stderrDiagnostic must not be null");
    Objects.requireNonNull(fallback, "fallback must not be null");
    try {
      CliPayloadOutput.write(stdout, fallback.payload());
    } catch (IOException stdoutFailure) {
      try {
        CliPayloadOutput.write(stderr, GridGrindCliJson.writeBytes(stderrDiagnostic, prettyJson));
      } catch (IOException stderrFailure) {
        stdoutFailure.addSuppressed(stderrFailure);
      }
      throw stdoutFailure;
    }
    try {
      CliPayloadOutput.write(stderr, GridGrindCliJson.writeBytes(stderrDiagnostic, prettyJson));
    } catch (IOException ignored) {
      // The stdout fallback payload is the primary recovery channel once response-file writing has
      // already failed. If stderr cannot mirror the transport diagnostic afterwards, keep the
      // successfully recovered stdout payload rather than re-failing the whole command.
      return;
    }
  }

  /** Value object for a stdout fallback payload. */
  static final class StdoutFallback {
    private final byte[] payload;

    private StdoutFallback(byte[] payload) {
      this.payload = Objects.requireNonNull(payload, "payload must not be null").clone();
    }

    byte[] payload() {
      return payload.clone();
    }
  }
}
