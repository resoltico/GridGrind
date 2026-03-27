package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.protocol.GridGrindJson;
import dev.erst.gridgrind.protocol.GridGrindRequest;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** Reads a GridGrind protocol request from stdin or an explicit request file path. */
final class CliRequestReader {
  /** Reads the request from stdin when the path is null, otherwise from the given file path. */
  GridGrindRequest read(Path requestPath, InputStream stdin) throws IOException {
    Objects.requireNonNull(stdin, "stdin must not be null");
    if (requestPath == null) {
      return GridGrindJson.readRequest(stdin);
    }
    try (InputStream requestInput = Files.newInputStream(requestPath)) {
      return GridGrindJson.readRequest(requestInput);
    }
  }
}
