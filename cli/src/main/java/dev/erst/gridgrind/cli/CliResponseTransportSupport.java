package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.cli.discovery.CliTransport;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.RequestDiagnosticRedactor;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.AccessDeniedException;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;

/** Shared response-file transport, fallback, and output-redaction mechanics. */
final class CliResponseTransportSupport {
  private CliResponseTransportSupport() {}

  static Path responseTargetPath(Path responsePath) {
    return responsePath.toAbsolutePath();
  }

  static void writePayload(Path targetPath, byte[] payload) throws IOException {
    Files.createDirectories(
        java.util.Objects.requireNonNull(
            targetPath.getParent(), "responsePath must not be a filesystem root"));
    try (OutputStream responseOutput =
        Files.newOutputStream(
            targetPath,
            java.nio.file.StandardOpenOption.CREATE_NEW,
            java.nio.file.StandardOpenOption.WRITE)) {
      writePayload(responseOutput, payload);
    }
  }

  static void writePayload(OutputStream outputStream, byte[] payload) throws IOException {
    CliPayloadOutput.write(outputStream, payload);
  }

  static void writeCliDiagnosticToStderr(
      OutputStream stderr,
      CliDiagnostic diagnostic,
      Optional<RequestDiagnosticRedactor> redactor,
      boolean prettyJson)
      throws IOException {
    java.util.Objects.requireNonNull(stderr, "stderr must not be null");
    java.util.Objects.requireNonNull(diagnostic, "diagnostic must not be null");
    writePayload(stderr, diagnosticBytes(diagnostic, redactor, prettyJson));
  }

  static byte[] diagnosticBytes(
      CliDiagnostic diagnostic, Optional<RequestDiagnosticRedactor> redactor, boolean prettyJson)
      throws IOException {
    return redact(redactor, GridGrindCliJson.writeBytes(diagnostic, prettyJson), prettyJson);
  }

  static byte[] redact(
      Optional<RequestDiagnosticRedactor> redactor, byte[] payload, boolean prettyJson)
      throws IOException {
    return redactor.isEmpty()
        ? payload
        : redactor.orElseThrow().redactSerializedJson(payload, prettyJson);
  }

  static GridGrindProblemDetail.Problem writeResponseProblem(
      IOException exception, Path targetPath) {
    var context =
        new ProblemContext.WriteResponse(
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
    java.util.Objects.requireNonNull(exception, "exception must not be null");
    java.util.Objects.requireNonNull(targetPath, "targetPath must not be null");
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
        diagnostic.problems(),
        Optional.of(java.util.Objects.requireNonNull(transport, "transport must not be null")));
  }

  static int exitCodeFor(GridGrindResponse response) {
    return switch (response) {
      case GridGrindResponse.Success _ -> 0;
      case GridGrindResponse.Failure _ -> 1;
    };
  }

  static int doctorExitCodeFor(RequestDoctorReport report) {
    java.util.Objects.requireNonNull(report, "report must not be null");
    return report.valid() ? 0 : 1;
  }

  private static Optional<String> fileSystemReason(FileSystemException exception) {
    String reason = exception.getReason();
    if (reason != null && !reason.isBlank()) {
      return Optional.of(reason);
    }
    String otherFile = exception.getOtherFile();
    if (otherFile != null && !otherFile.isBlank()) {
      return Optional.of("conflict with " + otherFile);
    }
    return Optional.empty();
  }
}
