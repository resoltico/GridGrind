package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.cli.discovery.CliTransport;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.AccessDeniedException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
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
    return writeCliDiagnosticNamed(responsePath, stdout, stderr, diagnostic, prettyJson);
  }

  /** Writes a request-content diagnostic and mirrors it to stderr when one file captures it. */
  int writeRequestDiagnostic(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliDiagnostic diagnostic,
      boolean prettyJson)
      throws IOException {
    return writeCliDiagnosticNamed(responsePath, stdout, stderr, diagnostic, prettyJson);
  }

  private int writeCliDiagnosticNamed(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliDiagnostic diagnostic,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(diagnostic, "diagnostic must not be null");
    if (responsePath.isEmpty()) {
      writePayload(stderr, GridGrindCliJson.writeBytes(diagnostic, prettyJson));
      return diagnostic.exitCode();
    }

    Path targetPath = responseTargetPath(responsePath.orElseThrow());
    try {
      CliDiagnostic persistedDiagnostic =
          diagnosticWithTransport(diagnostic, CliTransport.responseFile(targetPath.toString()));
      byte[] reportBytes = GridGrindCliJson.writeBytes(persistedDiagnostic, prettyJson);
      writePayload(targetPath, reportBytes);
      writeCliDiagnosticToStderr(stderr, persistedDiagnostic, prettyJson);
      return diagnostic.exitCode();
    } catch (IOException exception) {
      CliDiagnostic stdoutDiagnostic =
          diagnosticWithTransport(diagnostic, CliTransport.standardOutput());
      CliStdoutFallbackSupport.write(
          stderr,
          stdout,
          stdoutDiagnostic,
          CliStdoutFallbackSupport.cliDiagnostic(stdoutDiagnostic, prettyJson),
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
      writePayload(stdout, payload);
      return successExitCode;
    }

    Path targetPath = responseTargetPath(responsePath.orElseThrow());
    try {
      writePayload(targetPath, payload);
      return successExitCode;
    } catch (IOException exception) {
      CliDiagnostic stdoutDiagnostic =
          diagnosticWithTransport(
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
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(response, "response must not be null");
    if (responsePath.isEmpty()) {
      write(stdout, response, prettyJson);
      return logicalExitCode;
    }

    Path targetPath = responseTargetPath(responsePath.orElseThrow());
    try {
      writePayload(targetPath, GridGrindJsonOutput.writeResponseBytes(response, prettyJson));
      if (logicalExitCode != 0 && response instanceof GridGrindResponse.Failure failure) {
        writeCliDiagnosticToStderr(
            stderr,
            diagnosticWithTransport(
                CliDiagnostics.problemDiagnostic(logicalExitCode, "execute", failure.problem()),
                CliTransport.responseFile(targetPath.toString())),
            prettyJson);
      }
      return logicalExitCode;
    } catch (IOException exception) {
      GridGrindProblemDetail.Problem problem = writeResponseProblem(exception, targetPath);
      if (response instanceof GridGrindResponse.Failure failure) {
        problem =
            GridGrindProblems.appendCause(
                problem, GridGrindProblems.problemCause(failure.problem()));
      }
      GridGrindResponse.Failure fallbackResponse =
          GridGrindResponses.failure(
              GridGrindProtocolVersion.current(), response.persistence(), problem);
      CliDiagnostic stderrDiagnostic =
          diagnosticWithTransport(
              CliDiagnostics.problemDiagnostic(1, "execute", fallbackResponse.problem()),
              CliTransport.standardOutput());
      CliStdoutFallbackSupport.write(
          stderr,
          stdout,
          stderrDiagnostic,
          CliStdoutFallbackSupport.response(fallbackResponse, prettyJson),
          prettyJson);
      return 1;
    }
  }

  /** Writes one response to an already-open output stream, preserving caller stream ownership. */
  void write(OutputStream outputStream, GridGrindResponse response, boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(response, "response must not be null");
    writePayload(outputStream, GridGrindJsonOutput.writeResponseBytes(response, prettyJson));
  }

  /** Returns the process exit code associated with the response shape. */
  static int exitCodeFor(GridGrindResponse response) {
    return switch (response) {
      case GridGrindResponse.Success _ -> 0;
      case GridGrindResponse.Failure _ -> 1;
    };
  }

  /** Writes one doctor report to stdout or a configured response file. */
  int writeDoctorReport(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      RequestDoctorReport report,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(report, "report must not be null");
    if (responsePath.isEmpty()) {
      writeDoctorReport(stdout, report, prettyJson);
      return doctorExitCodeFor(report);
    }

    int exitCode = doctorExitCodeFor(report);
    Path targetPath = responseTargetPath(responsePath.orElseThrow());
    try {
      writePayload(
          targetPath, GridGrindJsonOutput.writeRequestDoctorReportBytes(report, prettyJson));
      if (report.valid()) {
        return exitCode;
      }
      GridGrindProblemDetail.Problem primaryProblem = report.primaryProblem().orElseThrow();
      writeCliDiagnosticToStderr(
          stderr,
          diagnosticWithTransport(
              CliDiagnostics.problemDiagnostic(exitCode, "doctor-request", primaryProblem),
              CliTransport.responseFile(targetPath.toString())),
          prettyJson);
      return exitCode;
    } catch (IOException exception) {
      GridGrindProblemDetail.Problem problem = writeResponseProblem(exception, targetPath);
      if (report.primaryProblem().isPresent()) {
        problem =
            GridGrindProblems.appendCause(
                problem, GridGrindProblems.problemCause(report.primaryProblem().orElseThrow()));
      }
      RequestDoctorReport fallbackReport =
          RequestDoctorReport.invalid(report.summary(), report.warnings(), problem);
      CliDiagnostic stderrDiagnostic =
          diagnosticWithTransport(
              CliDiagnostics.problemDiagnostic(
                  1, "doctor-request", fallbackReport.primaryProblem().orElseThrow()),
              CliTransport.standardOutput());
      CliStdoutFallbackSupport.write(
          stderr,
          stdout,
          stderrDiagnostic,
          CliStdoutFallbackSupport.doctorReport(fallbackReport, prettyJson),
          prettyJson);
      return 1;
    }
  }

  /** Writes one doctor report to an already-open output stream, preserving caller ownership. */
  void writeDoctorReport(OutputStream outputStream, RequestDoctorReport report, boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(report, "report must not be null");
    writePayload(
        outputStream, GridGrindJsonOutput.writeRequestDoctorReportBytes(report, prettyJson));
  }

  /** Returns the process exit code associated with one request doctor report. */
  static int doctorExitCodeFor(RequestDoctorReport report) {
    Objects.requireNonNull(report, "report must not be null");
    return report.valid() ? 0 : 1;
  }

  private static Path responseTargetPath(Path responsePath) {
    return responsePath.toAbsolutePath();
  }

  private static void writePayload(Path targetPath, byte[] payload) throws IOException {
    Files.createDirectories(
        Objects.requireNonNull(
            targetPath.getParent(), "responsePath must not be a filesystem root"));
    try (OutputStream responseOutput =
        Files.newOutputStream(
            targetPath,
            java.nio.file.StandardOpenOption.CREATE_NEW,
            java.nio.file.StandardOpenOption.WRITE)) {
      writePayload(responseOutput, payload);
    }
  }

  private static void writePayload(OutputStream outputStream, byte[] payload) throws IOException {
    CliPayloadOutput.write(outputStream, payload);
  }

  private static void writeCliDiagnosticToStderr(
      OutputStream stderr, CliDiagnostic diagnostic, boolean prettyJson) throws IOException {
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(diagnostic, "diagnostic must not be null");
    writePayload(stderr, GridGrindCliJson.writeBytes(diagnostic, prettyJson));
  }

  static GridGrindProblemDetail.Problem writeResponseProblem(
      IOException exception, Path targetPath) {
    var context =
        new dev.erst.gridgrind.contract.dto.ProblemContext.WriteResponse(
            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.ResponseOutput
                .responseFile(targetPath.toString()));
    String message = responseWriteMessage(exception, targetPath);
    return GridGrindProblems.problem(
        GridGrindProblemCode.IO_ERROR,
        message,
        context,
        java.util.List.of(
            new GridGrindProblemDetail.ProblemCause(
                GridGrindProblemCode.IO_ERROR, message, context.stage())));
  }

  static String responseWriteMessage(IOException exception, Path targetPath) {
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    return switch (exception) {
      case AccessDeniedException _ ->
          "Could not write response file " + targetPath + ": permission denied";
      case FileAlreadyExistsException _ when Files.isDirectory(targetPath) ->
          "Could not write response file " + targetPath + ": Is a directory";
      case FileAlreadyExistsException _ ->
          "Could not write response file "
              + targetPath
              + ": already exists; GridGrind never replaces an existing response file implicitly";
      case FileSystemException fileSystemException ->
          fileSystemReason(fileSystemException)
              .map(reason -> "Could not write response file " + targetPath + ": " + reason)
              .orElse("Could not write response file " + targetPath);
      default -> {
        String message = exception.getMessage();
        yield message == null || message.isBlank()
            ? "Could not write response file " + targetPath
            : "Could not write response file " + targetPath + ": " + message;
      }
    };
  }

  static CliDiagnostic diagnosticWithTransport(CliDiagnostic diagnostic, CliTransport transport) {
    return new CliDiagnostic(
        diagnostic.protocolVersion(),
        diagnostic.exitCode(),
        diagnostic.command(),
        diagnostic.suggestions(),
        diagnostic.problem(),
        Optional.of(Objects.requireNonNull(transport, "transport must not be null")));
  }

  private static java.util.Optional<String> fileSystemReason(FileSystemException exception) {
    String reason = exception.getReason();
    if (reason != null && !reason.isBlank()) {
      return java.util.Optional.of(reason);
    }
    String otherFile = exception.getOtherFile();
    if (otherFile != null && !otherFile.isBlank()) {
      return java.util.Optional.of("conflict with " + otherFile);
    }
    return java.util.Optional.empty();
  }
}
