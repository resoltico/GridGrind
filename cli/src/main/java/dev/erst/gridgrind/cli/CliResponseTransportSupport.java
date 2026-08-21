package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.json.RequestDiagnosticRedactor;
import java.io.IOException;
import java.io.OutputStream;
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
    writePayload(targetPath, payload, CliResponseTransportSupport::writeTemporaryPayload);
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

  static void writeTransportNoticeToStderr(OutputStream stderr, CliTransportNotice notice)
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
}
