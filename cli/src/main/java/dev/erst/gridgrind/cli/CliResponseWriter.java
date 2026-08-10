package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import dev.erst.gridgrind.contract.json.RequestDiagnosticRedactor;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Writes GridGrind responses to stdout or an explicit response file with a stdout fallback. */
final class CliResponseWriter {
  /** Writes one rejected-command result to stdout or the requested response file. */
  int writeCommandError(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CommandError commandError,
      boolean prettyJson)
      throws IOException {
    return writeCommandError(
        responsePath, stdout, stderr, commandError, Optional.empty(), prettyJson);
  }

  /** Writes one rejected-command result after applying the request-scoped secret boundary. */
  int writeCommandError(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CommandError commandError,
      RequestDiagnosticRedactor redactor,
      boolean prettyJson)
      throws IOException {
    return writeCommandError(
        responsePath,
        stdout,
        stderr,
        commandError,
        Optional.of(Objects.requireNonNull(redactor, "redactor must not be null")),
        prettyJson);
  }

  private int writeCommandError(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CommandError commandError,
      Optional<RequestDiagnosticRedactor> redactor,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(commandError, "commandError must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    return routePrimaryPayload(
        responsePath,
        stdout,
        stderr,
        CliResponseTransportSupport.commandErrorBytes(commandError, redactor, prettyJson),
        CliExitCodes.forCommandError(commandError));
  }

  /**
   * Writes one arbitrary command payload to stdout or a configured response file while also
   * reporting response-file fallback details on stderr.
   *
   * <p>When the response file cannot be written and stdout is writable, the already-rendered
   * primary payload is emitted there unchanged. An unavailable stdout transport is allowed to fail
   * closed rather than moving a primary payload to stderr.
   */
  int writePayload(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      byte[] payload,
      int successExitCode)
      throws IOException {
    return routePrimaryPayload(responsePath, stdout, stderr, payload, successExitCode);
  }

  /** Writes the response and returns one caller-chosen logical exit code on success. */
  int write(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      WorkbookResult response,
      int logicalExitCode,
      boolean prettyJson)
      throws IOException {
    return write(
        responsePath, stdout, stderr, response, logicalExitCode, Optional.empty(), prettyJson);
  }

  /** Writes one execution response after applying the request-scoped secret boundary. */
  int write(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      WorkbookResult response,
      int logicalExitCode,
      RequestDiagnosticRedactor redactor,
      boolean prettyJson)
      throws IOException {
    return write(
        responsePath,
        stdout,
        stderr,
        response,
        logicalExitCode,
        Optional.of(Objects.requireNonNull(redactor, "redactor must not be null")),
        prettyJson);
  }

  private int write(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      WorkbookResult response,
      int logicalExitCode,
      Optional<RequestDiagnosticRedactor> redactor,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(response, "response must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    return routePrimaryPayload(
        responsePath,
        stdout,
        stderr,
        CliResponseTransportSupport.redact(
            redactor,
            GridGrindJsonOutput.writeWorkbookResultBytes(response, prettyJson),
            prettyJson),
        logicalExitCode);
  }

  /** Writes one doctor report to stdout or a configured response file. */
  int writeDoctorReport(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      RequestDoctorReport report,
      boolean prettyJson)
      throws IOException {
    return writeDoctorReport(responsePath, stdout, stderr, report, Optional.empty(), prettyJson);
  }

  /** Writes one doctor report after applying the request-scoped secret boundary. */
  int writeDoctorReport(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      RequestDoctorReport report,
      RequestDiagnosticRedactor redactor,
      boolean prettyJson)
      throws IOException {
    return writeDoctorReport(
        responsePath,
        stdout,
        stderr,
        report,
        Optional.of(Objects.requireNonNull(redactor, "redactor must not be null")),
        prettyJson);
  }

  private int writeDoctorReport(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      RequestDoctorReport report,
      Optional<RequestDiagnosticRedactor> redactor,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(report, "report must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    return routePrimaryPayload(
        responsePath,
        stdout,
        stderr,
        CliResponseTransportSupport.redact(
            redactor,
            GridGrindJsonOutput.writeRequestDoctorReportBytes(report, prettyJson),
            prettyJson),
        CliExitCodes.forDoctorReport(report));
  }

  private static int routePrimaryPayload(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      byte[] payload,
      int logicalExitCode)
      throws IOException {
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(payload, "payload must not be null");
    if (responsePath.isEmpty()) {
      writePrimaryPayload(stdout, payload);
      return logicalExitCode;
    }

    Path targetPath = CliResponseTransportSupport.responseTargetPath(responsePath.orElseThrow());
    try {
      CliResponseTransportSupport.writePayload(targetPath, payload);
      return logicalExitCode;
    } catch (IOException exception) {
      writeStdoutFallback(stdout, stderr, payload, targetPath);
      return 1;
    }
  }

  private static void writeStdoutFallback(
      OutputStream stdout, OutputStream stderr, byte[] payload, Path responsePath)
      throws IOException {
    writePrimaryPayload(stdout, payload);
    try {
      CliResponseTransportSupport.writeTransportNoticeToStderr(
          stderr, CliTransportNotice.stdoutFallback(responsePath.toString()));
    } catch (IOException ignored) {
      // The recovered stdout payload is primary once the requested response file cannot be used.
    }
  }

  private static void writePrimaryPayload(OutputStream stdout, byte[] payload) {
    try {
      CliResponseTransportSupport.writePayload(stdout, payload);
    } catch (IOException exception) {
      throw new CliPrimaryOutputException(exception);
    }
  }
}
