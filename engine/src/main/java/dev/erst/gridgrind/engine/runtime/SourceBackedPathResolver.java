package dev.erst.gridgrind.engine.runtime;

import java.nio.file.Files;
import java.nio.file.Path;

/** Resolves authored source-file paths against the execution working directory. */
final class SourceBackedPathResolver {
  private SourceBackedPathResolver() {}

  static Path resolvePath(String rawPath, Path workingDirectory, String inputKind)
      throws InputSourceReadException {
    try {
      Path resolved = ExecutionRequestPaths.normalizePath(rawPath, workingDirectory); // LIM-030
      if (Files.isDirectory(resolved)) {
        throw new InputSourceReadException(
            inputKind + " path must resolve to a file, not a directory: " + resolved,
            inputKind,
            resolved.toString(),
            null);
      }
      return resolved;
    } catch (IllegalArgumentException exception) {
      throw new InputSourceReadException(
          "Invalid " + inputKind + " path: " + rawPath, inputKind, rawPath, exception);
    }
  }
}
