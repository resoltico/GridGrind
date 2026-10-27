package dev.erst.gridgrind.cli.io;

import dev.erst.gridgrind.cli.command.CliExitCodes;
import dev.erst.gridgrind.cli.render.CliTransportNotice;
import dev.erst.gridgrind.cli.render.GridGrindCliJson;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.result.CommandError;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import dev.erst.gridgrind.json.codec.RequestDiagnosticRedactor;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Writes GridGrind responses to stdout or an explicit response file with a stdout fallback. */
public final class CliResponseWriter {
  /** Writes one rejected-command result to stdout or the requested response file. */
  public int writeCommandError(
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
  public int writeCommandError(
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
        commandErrorBytes(commandError, redactor, prettyJson),
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
  public int writePayload(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      byte[] payload,
      int successExitCode)
      throws IOException {
    return routePrimaryPayload(responsePath, stdout, stderr, payload, successExitCode);
  }

  /** Writes the response and returns one caller-chosen logical exit code on success. */
  public int write(
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

  /** Writes one execution payload through a destination reserved before workbook side effects. */
  public int writeReserved(
      CliResponseReservation reservation,
      OutputStream stdout,
      OutputStream stderr,
      WorkbookResult response,
      int logicalExitCode,
      RequestDiagnosticRedactor redactor,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(reservation, "reservation must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(response, "response must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    byte[] payload =
        redact(
            Optional.of(redactor),
            GridGrindJsonOutput.writeWorkbookResultBytes(response, prettyJson),
            prettyJson);
    try {
      reservation.write(payload);
      return logicalExitCode;
    } catch (IOException exception) {
      writeStdoutFallback(stdout, stderr, payload, reservation.responsePath());
      return 1;
    }
  }

  /** Writes one execution response after applying the request-scoped secret boundary. */
  public int write(
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
        redact(
            redactor,
            GridGrindJsonOutput.writeWorkbookResultBytes(response, prettyJson),
            prettyJson),
        logicalExitCode);
  }

  /** Writes one doctor report to stdout or a configured response file. */
  public int writeDoctorReport(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      RequestDoctorReport report,
      boolean prettyJson)
      throws IOException {
    return writeDoctorReport(responsePath, stdout, stderr, report, Optional.empty(), prettyJson);
  }

  /** Writes one doctor report after applying the request-scoped secret boundary. */
  public int writeDoctorReport(
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
        redact(
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

    Path targetPath = responseTargetPath(responsePath.orElseThrow());
    try {
      writePayload(targetPath, payload);
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
      writeTransportNoticeToStderr(
          stderr, CliTransportNotice.stdoutFallback(responsePath.toString()));
    } catch (IOException ignored) {
      // The recovered stdout payload is primary once the requested response file cannot be used.
    }
  }

  private static void writePrimaryPayload(OutputStream stdout, byte[] payload) {
    try {
      writePayload(stdout, payload);
    } catch (IOException exception) {
      throw new CliPrimaryOutputException(exception);
    }
  }

  static Path responseTargetPath(Path responsePath) {
    return responsePath.toAbsolutePath();
  }

  static void writePayload(Path targetPath, byte[] payload) throws IOException {
    writePayload(targetPath, payload, CliResponseWriter::writeTemporaryPayload);
  }

  static void writePayload(Path targetPath, byte[] payload, TemporaryPayloadWriter temporaryWriter)
      throws IOException {
    java.util.Objects.requireNonNull(targetPath, "targetPath must not be null");
    java.util.Objects.requireNonNull(payload, "payload must not be null");
    java.util.Objects.requireNonNull(temporaryWriter, "temporaryWriter must not be null");
    Path parent =
        java.util.Objects.requireNonNull(
            targetPath.getParent(), "responsePath must not be a filesystem root");
    Files.createDirectories(parent);
    Path temporaryPath = Files.createTempFile(parent, ".gridgrind-response-", ".tmp");
    try {
      temporaryWriter.write(temporaryPath, payload);
      Files.move(temporaryPath, targetPath);
    } catch (IOException | RuntimeException | Error exception) {
      deleteFailedStagingFile(temporaryPath, exception);
      throw exception;
    }
  }

  private static void deleteFailedStagingFile(Path temporaryPath, Throwable failure) {
    try {
      Files.deleteIfExists(temporaryPath);
    } catch (IOException cleanupFailure) {
      failure.addSuppressed(cleanupFailure);
    }
  }

  private static void writeTemporaryPayload(Path temporaryPath, byte[] payload) throws IOException {
    try (OutputStream responseOutput =
        Files.newOutputStream(
            temporaryPath,
            java.nio.file.StandardOpenOption.TRUNCATE_EXISTING,
            java.nio.file.StandardOpenOption.WRITE)) {
      writePayload(responseOutput, payload);
    }
  }

  static void writePayload(OutputStream outputStream, byte[] payload) throws IOException {
    CliPayloadOutput.write(outputStream, payload);
  }

  /** Writes one structured transport notice to the diagnostic stream. */
  public static void writeTransportNoticeToStderr(OutputStream stderr, CliTransportNotice notice)
      throws IOException {
    java.util.Objects.requireNonNull(stderr, "stderr must not be null");
    java.util.Objects.requireNonNull(notice, "notice must not be null");
    writePayload(stderr, GridGrindCliJson.writeBytes(notice));
  }

  static byte[] commandErrorBytes(
      CommandError commandError, Optional<RequestDiagnosticRedactor> redactor, boolean prettyJson)
      throws IOException {
    return redact(redactor, GridGrindCliJson.writeBytes(commandError, prettyJson), prettyJson);
  }

  static byte[] redact(
      Optional<RequestDiagnosticRedactor> redactor, byte[] payload, boolean prettyJson)
      throws IOException {
    return redactor.isEmpty()
        ? payload
        : redactor.orElseThrow().redactSerializedJson(payload, prettyJson);
  }

  /** Writes the complete response bytes into a staging file before publication. */
  @FunctionalInterface
  interface TemporaryPayloadWriter {
    /** Writes the complete payload into a staging file before it becomes externally visible. */
    void write(Path temporaryPath, byte[] payload) throws IOException;
  }

  /** Appends intake warnings while preserving the classified execution outcome. */
  public static WorkbookResult withTransportWarnings(
      WorkbookResult response, RequestIntake analysis) {
    return withTransportWarnings(response, analysis.warnings());
  }

  /** Appends intake warnings while preserving the classified execution outcome. */
  public static WorkbookResult withTransportWarnings(
      WorkbookResult response, List<RequestWarning> analysisWarnings) {
    if (analysisWarnings.isEmpty()) {
      return response;
    }
    List<RequestWarning> warnings = new ArrayList<>();
    return switch (response) {
      case WorkbookResult.Success success -> {
        warnings.addAll(success.warnings());
        warnings.addAll(analysisWarnings);
        yield new WorkbookResult.Success(
            success.protocolVersion(),
            success.planId(),
            success.journal(),
            success.calculation(),
            success.persistence(),
            success.evidence(),
            warnings,
            success.assertions(),
            success.inspections());
      }
      case WorkbookResult.Failure failure -> {
        warnings.addAll(failure.warnings());
        warnings.addAll(analysisWarnings);
        yield new WorkbookResult.Failure(
            failure.protocolVersion(),
            failure.planId(),
            failure.journal(),
            failure.calculation(),
            failure.persistence(),
            failure.evidence(),
            warnings,
            failure.assertions(),
            failure.inspections(),
            failure.problem());
      }
    };
  }

  /** Renders a command payload before taking ownership of its primary output. */
  public int writeRenderedPayload(
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      OutputRenderer renderer)
      throws IOException {
    java.io.ByteArrayOutputStream buffer = new java.io.ByteArrayOutputStream();
    renderer.write(buffer);
    return writePayload(responsePath, stdout, stderr, buffer.toByteArray(), 0);
  }

  /** Renders one command payload into a caller-owned buffer. */
  @FunctionalInterface
  public interface OutputRenderer {
    /** Writes the complete payload into the supplied stream. */
    void write(OutputStream outputStream) throws IOException;
  }
}
