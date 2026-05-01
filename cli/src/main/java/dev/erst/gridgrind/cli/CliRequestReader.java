package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.AccessDeniedException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.util.Objects;

/** Reads a GridGrind protocol request from stdin or an explicit request file path. */
final class CliRequestReader {
  /** Reads the request from stdin when the path is null, otherwise from the given file path. */
  WorkbookPlan read(Path requestPath, InputStream stdin) throws IOException {
    Objects.requireNonNull(stdin, "stdin must not be null");
    if (requestPath == null) {
      return GridGrindJson.readRequest(stdin);
    }
    Path normalizedRequestPath = requestPath.toAbsolutePath().normalize();
    validateReadableRequestPath(normalizedRequestPath);
    GridGrindJson.requireSupportedRequestLength(Files.size(normalizedRequestPath));
    try (InputStream requestInput = Files.newInputStream(normalizedRequestPath)) {
      return GridGrindJson.readRequest(requestInput);
    }
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
