package dev.erst.gridgrind.executor;

import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;

/** Resolves authored source-file paths against the execution working directory. */
final class SourceBackedPathResolver {
  private SourceBackedPathResolver() {}

  static Path resolvePath(String rawPath, Path workingDirectory, String inputKind)
      throws InputSourceReadException {
    try {
      Path candidate = Path.of(rawPath);
      Path resolved =
          candidate.isAbsolute()
              ? candidate.toAbsolutePath().normalize()
              : workingDirectory.resolve(candidate).normalize();
      if (Files.isDirectory(resolved)) {
        throw new InputSourceReadException(
            inputKind + " path must resolve to a file, not a directory: " + resolved,
            inputKind,
            resolved.toString(),
            null);
      }
      return resolved;
    } catch (InvalidPathException exception) {
      throw new InputSourceReadException(
          "Invalid " + inputKind + " path: " + rawPath, inputKind, rawPath, exception);
    }
  }
}
