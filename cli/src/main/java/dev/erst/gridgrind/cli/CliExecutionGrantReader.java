package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

/** Reads one explicit CLI host-grant document without treating it as request-owned plan input. */
final class CliExecutionGrantReader {
  private CliExecutionGrantReader() {}

  static GridGrindExecutionGrant read(Path grantPath, Path cliWorkingDirectory) throws IOException {
    Path normalizedGrantPath =
        Objects.requireNonNull(grantPath, "grantPath must not be null")
            .toAbsolutePath()
            .normalize();
    CliExecutionGrantDocument document =
        GridGrindCliJson.readBytes(
            Files.readAllBytes(normalizedGrantPath), CliExecutionGrantDocument.class);
    return document.toExecutionGrant(cliWorkingDirectory);
  }
}
