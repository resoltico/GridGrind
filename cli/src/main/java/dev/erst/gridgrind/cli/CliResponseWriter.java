package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.cli.discovery.CliTransport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import dev.erst.gridgrind.contract.json.RequestDiagnosticRedactor;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Writes GridGrind responses to stdout or an explicit response file with structured fallback. */
final class CliResponseWriter {
  /**
   * Writes a CLI diagnostic to the configured destination.
   *
   * <p>When {@code responsePath} is empty, CLI failure JSON goes to {@code stderr} so non-success
   * CLI transport failures do not masquerade as primary stdout payloads. When a response path is
   * present, the same CLI diagnostic is also mirrored to {@code stderr} with transport metadata so
   * stderr stays machine-readable while still naming the persisted file or stdout fallback.
   */
  int writeCliDiagnostic(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliDiagnostic diagnostic,
      boolean prettyJson)
      throws IOException {
    return writeCliDiagnosticNamed(
        responsePath, stdout, stderr, diagnostic, Optional.empty(), prettyJson);
  }

  /** Writes a request-content diagnostic and mirrors it to stderr when one file captures it. */
  int writeRequestDiagnostic(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliDiagnostic diagnostic,
      boolean prettyJson)
      throws IOException {
    return writeCliDiagnosticNamed(
        responsePath, stdout, stderr, diagnostic, Optional.empty(), prettyJson);
  }

  /** Writes one request diagnostic after applying the request-scoped secret boundary. */
  int writeRequestDiagnostic(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliDiagnostic diagnostic,
      RequestDiagnosticRedactor redactor,
      boolean prettyJson)
      throws IOException {
    return writeCliDiagnosticNamed(
        responsePath,
        stdout,
        stderr,
        diagnostic,
        Optional.of(Objects.requireNonNull(redactor, "redactor must not be null")),
        prettyJson);
  }

  private int writeCliDiagnosticNamed(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliDiagnostic diagnostic,
      Optional<RequestDiagnosticRedactor> redactor,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(diagnostic, "diagnostic must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    if (responsePath.isEmpty()) {
      CliResponseTransportSupport.writePayload(
          stderr, CliResponseTransportSupport.diagnosticBytes(diagnostic, redactor, prettyJson));
      return diagnostic.exitCode();
    }

    Path targetPath = CliResponseTransportSupport.responseTargetPath(responsePath.orElseThrow());
    try {
      CliDiagnostic persistedDiagnostic =
          CliResponseTransportSupport.diagnosticWithTransport(
              diagnostic, CliTransport.responseFile(targetPath.toString()));
      byte[] reportBytes =
          CliResponseTransportSupport.diagnosticBytes(persistedDiagnostic, redactor, prettyJson);
      CliResponseTransportSupport.writePayload(targetPath, reportBytes);
      CliResponseTransportSupport.writeCliDiagnosticToStderr(
          stderr, persistedDiagnostic, redactor, prettyJson);
      return diagnostic.exitCode();
    } catch (IOException exception) {
      CliDiagnostic stdoutDiagnostic =
          CliResponseTransportSupport.diagnosticWithTransport(
              diagnostic, CliTransport.standardOutput());
      CliStdoutFallbackSupport.write(
          stderr,
          stdout,
          stdoutDiagnostic,
          CliStdoutFallbackSupport.redacted(
              CliStdoutFallbackSupport.cliDiagnostic(stdoutDiagnostic, prettyJson),
              redactor,
              prettyJson),
          redactor,
          prettyJson);
      return diagnostic.exitCode();
    }
  }

  /**
   * Writes one arbitrary command payload to stdout or a configured response file while also
   * reporting response-file fallback details on stderr.
   *
   * <p>When the response file cannot be written, a structured failure response is emitted to stdout
   * so every command family keeps the same fallback contract.
   */
  int writePayload(
      String command,
      String payloadName,
      Optional<String> stdoutSuggestion,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      byte[] payload,
      int successExitCode,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(payloadName, "payloadName must not be null");
    Objects.requireNonNull(stdoutSuggestion, "stdoutSuggestion must not be null");
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(payload, "payload must not be null");
    if (responsePath.isEmpty()) {
      CliResponseTransportSupport.writePayload(stdout, payload);
      return successExitCode;
    }

    Path targetPath = CliResponseTransportSupport.responseTargetPath(responsePath.orElseThrow());
    try {
      CliResponseTransportSupport.writePayload(targetPath, payload);
      return successExitCode;
    } catch (IOException exception) {
      CliDiagnostic stdoutDiagnostic =
          CliResponseTransportSupport.diagnosticWithTransport(
              CliDiagnostics.responseWriteFailure(
                  command, payloadName, targetPath, exception, stdoutSuggestion),
              CliTransport.standardOutput());
      CliStdoutFallbackSupport.write(
          stderr,
          stdout,
          stdoutDiagnostic,
          CliStdoutFallbackSupport.cliDiagnostic(stdoutDiagnostic, prettyJson),
          prettyJson);
      return 1;
    }
  }

  /** Writes the response and returns one caller-chosen logical exit code on success. */
  int write(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      GridGrindResponse response,
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
      GridGrindResponse response,
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
      GridGrindResponse response,
      int logicalExitCode,
      Optional<RequestDiagnosticRedactor> redactor,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(response, "response must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    if (responsePath.isEmpty()) {
      CliResponseTransportSupport.writePayload(
          stdout,
          CliResponseTransportSupport.redact(
              redactor, GridGrindJsonOutput.writeResponseBytes(response, prettyJson), prettyJson));
      return logicalExitCode;
    }

    Path targetPath = CliResponseTransportSupport.responseTargetPath(responsePath.orElseThrow());
    try {
      CliResponseTransportSupport.writePayload(
          targetPath,
          CliResponseTransportSupport.redact(
              redactor, GridGrindJsonOutput.writeResponseBytes(response, prettyJson), prettyJson));
      if (logicalExitCode != 0 && response instanceof GridGrindResponse.Failure failure) {
        CliResponseTransportSupport.writeCliDiagnosticToStderr(
            stderr,
            CliResponseTransportSupport.diagnosticWithTransport(
                CliDiagnostics.problemDiagnostic(logicalExitCode, "execute", failure.problem()),
                CliTransport.responseFile(targetPath.toString())),
            redactor,
            prettyJson);
      }
      return logicalExitCode;
    } catch (IOException exception) {
      GridGrindProblemDetail.Problem problem =
          CliResponseTransportSupport.writeResponseProblem(exception, targetPath);
      if (response instanceof GridGrindResponse.Failure failure) {
        problem =
            GridGrindProblems.appendCause(
                problem, GridGrindProblems.problemCause(failure.problem()));
      }
      GridGrindResponse.Failure fallbackResponse =
          GridGrindResponses.failure(
              GridGrindProtocolVersion.current(), response.persistence(), problem);
      CliDiagnostic stderrDiagnostic =
          CliResponseTransportSupport.diagnosticWithTransport(
              CliDiagnostics.problemDiagnostic(1, "execute", fallbackResponse.problem()),
              CliTransport.standardOutput());
      CliStdoutFallbackSupport.write(
          stderr,
          stdout,
          stderrDiagnostic,
          CliStdoutFallbackSupport.redacted(
              CliStdoutFallbackSupport.response(fallbackResponse, prettyJson),
              redactor,
              prettyJson),
          redactor,
          prettyJson);
      return 1;
    }
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
    if (responsePath.isEmpty()) {
      CliResponseTransportSupport.writePayload(
          stdout,
          CliResponseTransportSupport.redact(
              redactor,
              GridGrindJsonOutput.writeRequestDoctorReportBytes(report, prettyJson),
              prettyJson));
      return CliResponseTransportSupport.doctorExitCodeFor(report);
    }

    int exitCode = CliResponseTransportSupport.doctorExitCodeFor(report);
    Path targetPath = CliResponseTransportSupport.responseTargetPath(responsePath.orElseThrow());
    try {
      CliResponseTransportSupport.writePayload(
          targetPath,
          CliResponseTransportSupport.redact(
              redactor,
              GridGrindJsonOutput.writeRequestDoctorReportBytes(report, prettyJson),
              prettyJson));
      if (report.valid()) {
        return exitCode;
      }
      GridGrindProblemDetail.Problem primaryProblem = report.primaryProblem().orElseThrow();
      CliResponseTransportSupport.writeCliDiagnosticToStderr(
          stderr,
          CliResponseTransportSupport.diagnosticWithTransport(
              CliDiagnostics.problemDiagnostic(exitCode, "doctor-request", primaryProblem),
              CliTransport.responseFile(targetPath.toString())),
          redactor,
          prettyJson);
      return exitCode;
    } catch (IOException exception) {
      GridGrindProblemDetail.Problem problem =
          CliResponseTransportSupport.writeResponseProblem(exception, targetPath);
      if (report.primaryProblem().isPresent()) {
        problem =
            GridGrindProblems.appendCause(
                problem, GridGrindProblems.problemCause(report.primaryProblem().orElseThrow()));
      }
      RequestDoctorReport fallbackReport =
          RequestDoctorReport.invalid(report.summary(), report.warnings(), problem);
      CliDiagnostic stderrDiagnostic =
          CliResponseTransportSupport.diagnosticWithTransport(
              CliDiagnostics.problemDiagnostic(
                  1, "doctor-request", fallbackReport.primaryProblem().orElseThrow()),
              CliTransport.standardOutput());
      CliStdoutFallbackSupport.write(
          stderr,
          stdout,
          stderrDiagnostic,
          CliStdoutFallbackSupport.redacted(
              CliStdoutFallbackSupport.doctorReport(fallbackReport, prettyJson),
              redactor,
              prettyJson),
          redactor,
          prettyJson);
      return 1;
    }
  }
}
