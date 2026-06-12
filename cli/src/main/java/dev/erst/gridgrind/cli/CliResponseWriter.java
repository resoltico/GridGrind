package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.AccessDeniedException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Writes GridGrind responses to stdout or an explicit response file with structured fallback. */
final class CliResponseWriter {
  /**
   * Writes a CLI failure report to the configured destination.
   *
   * <p>When {@code responsePath} is empty, failure JSON goes to {@code stderr} so success payloads
   * remain the only stdout traffic. The {@code stderr} stream is also used in the response-file
   * path: to write a human-readable file pointer after a successful write, or to emit a fallback
   * notice when the file write itself fails.
   */
  int writeCliFailureReport(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliFailureReport report)
      throws IOException {
    return writeCliFailureReportNamed("CLI failure report", responsePath, stdout, stderr, report);
  }

  /**
   * Writes a request-content failure report, emitting a "request failure report" stderr pointer.
   */
  int writeRequestFailureReport(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliFailureReport report)
      throws IOException {
    return writeCliFailureReportNamed(
        "request failure report", responsePath, stdout, stderr, report);
  }

  private int writeCliFailureReportNamed(
      String payloadName,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliFailureReport report)
      throws IOException {
    Objects.requireNonNull(report, "report must not be null");
    if (responsePath.isEmpty()) {
      writePayload(stderr, GridGrindCliJson.writeCliFailureReportBytes(report));
      return report.exitCode();
    }

    Path targetPath = responseTargetPath(responsePath.orElseThrow());
    try {
      writePayload(targetPath, GridGrindCliJson.writeCliFailureReportBytes(report));
      writeNonSuccessPointerIfNeeded(
          stderr,
          report.exitCode(),
          targetPath,
          payloadName,
          "failure",
          Optional.of(report.code().name() + ": " + report.message()));
      return report.exitCode();
    } catch (IOException exception) {
      writeStdoutPayloadFallbackNotice(stderr, exception, targetPath, payloadName);
      writePayload(stdout, GridGrindCliJson.writeCliFailureReportBytes(report));
      return report.exitCode();
    }
  }

  /**
   * Writes one arbitrary command payload to stdout or a configured response file.
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
      byte[] payload,
      int successExitCode)
      throws IOException {
    return writePayload(
        command,
        payloadName,
        stdoutSuggestion,
        responsePath,
        stdout,
        OutputStream.nullOutputStream(),
        payload,
        successExitCode);
  }

  /**
   * Writes one arbitrary command payload to stdout or a configured response file while also
   * reporting response-file fallback details on stderr.
   */
  int writePayload(
      String command,
      String payloadName,
      Optional<String> stdoutSuggestion,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      byte[] payload,
      int successExitCode)
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
      writeStdoutPayloadFallbackNotice(stderr, exception, targetPath, payloadName);
      writePayload(
          stdout,
          GridGrindCliJson.writeCliFailureReportBytes(
              CliFailureReports.responseWriteFailure(
                  command, payloadName, targetPath, exception, stdoutSuggestion)));
      return 1;
    }
  }

  /** Writes the response to the configured destination and returns the corresponding exit code. */
  int write(Optional<Path> responsePath, OutputStream stdout, GridGrindResponse response)
      throws IOException {
    return write(responsePath, stdout, OutputStream.nullOutputStream(), response);
  }

  /** Writes the response and returns one caller-chosen logical exit code on success. */
  int write(
      Optional<Path> responsePath,
      OutputStream stdout,
      GridGrindResponse response,
      int logicalExitCode)
      throws IOException {
    return write(responsePath, stdout, OutputStream.nullOutputStream(), response, logicalExitCode);
  }

  /** Writes the response to the configured destination and returns the corresponding exit code. */
  int write(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      GridGrindResponse response)
      throws IOException {
    return write(responsePath, stdout, stderr, response, exitCodeFor(response));
  }

  /** Writes the response and returns one caller-chosen logical exit code on success. */
  int write(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      GridGrindResponse response,
      int logicalExitCode)
      throws IOException {
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(response, "response must not be null");
    if (responsePath.isEmpty()) {
      write(stdout, response);
      return logicalExitCode;
    }

    Path targetPath = responseTargetPath(responsePath.orElseThrow());
    try {
      writePayload(targetPath, GridGrindJson.writeResponseBytes(response));
      writeNonSuccessPointerIfNeeded(
          stderr,
          logicalExitCode,
          targetPath,
          "response",
          "failure",
          response instanceof GridGrindResponse.Failure failure
              ? Optional.of(failure.problem().code().name() + ": " + failure.problem().message())
              : Optional.empty());
      return logicalExitCode;
    } catch (IOException exception) {
      writeStdoutPayloadFallbackNotice(stderr, exception, targetPath, "response");
      GridGrindProblemDetail.Problem problem = writeResponseProblem(exception, targetPath);
      if (response instanceof GridGrindResponse.Failure failure) {
        problem =
            GridGrindProblems.appendCause(
                problem, GridGrindProblems.problemCause(failure.problem()));
      }
      write(stdout, GridGrindResponses.failure(GridGrindProtocolVersion.current(), problem));
      return 1;
    }
  }

  /** Writes one response to an already-open output stream, preserving caller stream ownership. */
  void write(OutputStream outputStream, GridGrindResponse response) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(response, "response must not be null");
    writePayload(outputStream, GridGrindJson.writeResponseBytes(response));
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
      Optional<Path> responsePath, OutputStream stdout, RequestDoctorReport report)
      throws IOException {
    return writeDoctorReport(responsePath, stdout, OutputStream.nullOutputStream(), report);
  }

  /** Writes one doctor report to stdout or a configured response file. */
  int writeDoctorReport(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      RequestDoctorReport report)
      throws IOException {
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(report, "report must not be null");
    if (responsePath.isEmpty()) {
      writeDoctorReport(stdout, report);
      return doctorExitCodeFor(report);
    }

    Path targetPath = responseTargetPath(responsePath.orElseThrow());
    try {
      writePayload(targetPath, GridGrindJson.writeRequestDoctorReportBytes(report));
      writeNonSuccessPointerIfNeeded(
          stderr,
          doctorExitCodeFor(report),
          targetPath,
          "doctor report",
          "problems",
          report.primaryProblem().map(problem -> problem.code().name() + ": " + problem.message()));
      return doctorExitCodeFor(report);
    } catch (IOException exception) {
      writeStdoutPayloadFallbackNotice(stderr, exception, targetPath, "doctor report");
      GridGrindProblemDetail.Problem problem = writeResponseProblem(exception, targetPath);
      if (report.primaryProblem().isPresent()) {
        problem =
            GridGrindProblems.appendCause(
                problem, GridGrindProblems.problemCause(report.primaryProblem().orElseThrow()));
      }
      writeDoctorReport(
          stdout, RequestDoctorReport.invalid(report.summary(), report.warnings(), problem));
      return 1;
    }
  }

  /** Writes one doctor report to an already-open output stream, preserving caller ownership. */
  void writeDoctorReport(OutputStream outputStream, RequestDoctorReport report) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(report, "report must not be null");
    writePayload(outputStream, GridGrindJson.writeRequestDoctorReportBytes(report));
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
    try (OutputStream responseOutput = Files.newOutputStream(targetPath)) {
      writePayload(responseOutput, payload);
    }
  }

  private static void writePayload(OutputStream outputStream, byte[] payload) throws IOException {
    CliPayloadOutput.write(outputStream, payload);
  }

  private static void writeNonSuccessPointerIfNeeded(
      OutputStream stderr,
      int exitCode,
      Path targetPath,
      String payloadName,
      String problemNoun,
      Optional<String> problemSummary)
      throws IOException {
    if (exitCode == 0) {
      return;
    }
    String summary = formattedProblemSummary(problemSummary);
    String line =
        "GridGrind wrote the "
            + payloadName
            + " to "
            + targetPath
            + "; inspect that file for "
            + problemNoun
            + summary
            + '.'
            + System.lineSeparator();
    stderr.write(line.getBytes(StandardCharsets.UTF_8));
    stderr.flush();
  }

  static String formattedProblemSummary(Optional<String> problemSummary) {
    Objects.requireNonNull(problemSummary, "problemSummary must not be null");
    return problemSummary.filter(text -> !text.isBlank()).map(text -> " [" + text + "]").orElse("");
  }

  private static void writeStdoutPayloadFallbackNotice(
      OutputStream stderr, IOException exception, Path targetPath, String payloadName)
      throws IOException {
    String line =
        responseWriteMessage(exception, targetPath)
            + ". Wrote the "
            + payloadName
            + " to stdout instead."
            + System.lineSeparator();
    stderr.write(line.getBytes(StandardCharsets.UTF_8));
    stderr.flush();
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
