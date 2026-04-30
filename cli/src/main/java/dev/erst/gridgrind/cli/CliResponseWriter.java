package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.executor.GridGrindProblems;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** Writes GridGrind responses to stdout or an explicit response file with structured fallback. */
final class CliResponseWriter {
  /**
   * Writes one arbitrary command payload to stdout or a configured response file.
   *
   * <p>When the response file cannot be written, a structured failure response is emitted to stdout
   * so every command family keeps the same fallback contract.
   */
  int writePayload(Path responsePath, OutputStream stdout, byte[] payload, int successExitCode)
      throws IOException {
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(payload, "payload must not be null");
    if (responsePath == null) {
      writePayload(stdout, payload);
      return successExitCode;
    }

    Path targetPath = responseTargetPath(responsePath);
    try {
      writePayload(targetPath, payload);
      return successExitCode;
    } catch (IOException exception) {
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.WriteResponse(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.ResponseOutput
                      .responseFile(targetPath.toString())));
      write(stdout, GridGrindResponses.failure(GridGrindProtocolVersion.current(), problem));
      return 1;
    }
  }

  /** Writes the response to the configured destination and returns the corresponding exit code. */
  int write(Path responsePath, OutputStream stdout, GridGrindResponse response) throws IOException {
    return write(responsePath, stdout, response, exitCodeFor(response));
  }

  /** Writes the response and returns one caller-chosen logical exit code on success. */
  int write(Path responsePath, OutputStream stdout, GridGrindResponse response, int logicalExitCode)
      throws IOException {
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(response, "response must not be null");
    if (responsePath == null) {
      write(stdout, response);
      return logicalExitCode;
    }

    Path targetPath = responseTargetPath(responsePath);
    try {
      writePayload(targetPath, GridGrindJson.writeResponseBytes(response));
      return logicalExitCode;
    } catch (IOException exception) {
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.WriteResponse(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.ResponseOutput
                      .responseFile(targetPath.toString())));
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
  int writeDoctorReport(Path responsePath, OutputStream stdout, RequestDoctorReport report)
      throws IOException {
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(report, "report must not be null");
    if (responsePath == null) {
      writeDoctorReport(stdout, report);
      return doctorExitCodeFor(report);
    }

    Path targetPath = responseTargetPath(responsePath);
    try {
      writePayload(targetPath, GridGrindJson.writeRequestDoctorReportBytes(report));
      return doctorExitCodeFor(report);
    } catch (IOException exception) {
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.WriteResponse(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.ResponseOutput
                      .responseFile(targetPath.toString())));
      if (report.problem().isPresent()) {
        problem =
            GridGrindProblems.appendCause(
                problem, GridGrindProblems.problemCause(report.problem().orElseThrow()));
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
    outputStream.write(payload);
    if (payload.length == 0 || payload[payload.length - 1] != '\n') {
      outputStream.write('\n');
    }
    outputStream.flush();
  }
}
