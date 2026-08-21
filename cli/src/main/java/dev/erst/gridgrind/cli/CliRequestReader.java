package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.AccessDeniedException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Reads a GridGrind protocol request from stdin or an explicit request file path. */
final class CliRequestReader {
  /** Reads raw request bytes from stdin or an explicit request file path. */
  byte[] readBytes(Optional<Path> requestPath, InputStream stdin) throws IOException {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    if (requestPath.isEmpty() || CliPathArguments.isStandardInputPath(requestPath)) {
      byte[] requestBytes = stdin.readAllBytes();
      GridGrindJson.requireSupportedRequestLength(requestBytes.length);
      return requestBytes;
    }
    Path normalizedRequestPath = requestPath.orElseThrow().toAbsolutePath().normalize();
    validateReadableRequestPath(normalizedRequestPath);
    GridGrindJson.requireSupportedRequestLength(Files.size(normalizedRequestPath));
    return Files.readAllBytes(normalizedRequestPath);
  }

  private static void validateReadableRequestPath(Path requestPath) throws IOException {
    if (!Files.exists(requestPath)) {
      throw new NoSuchFileException("Request file not found: " + requestPath);
    }
    if (!Files.isRegularFile(requestPath)) {
      throw new IOException("Request path is not a regular file: " + requestPath);
    }
    if (!Files.isReadable(requestPath)) {
      throw new AccessDeniedException("Request file is not readable: " + requestPath);
    }
  }
}
