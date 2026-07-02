package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Path-typed CLI argument helpers shared by parsing and fallback reporting. */
final class CliPathArguments {
  private static final String STANDARD_INPUT_PATH = "-";

  private CliPathArguments() {}

  static Optional<Path> responsePath(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    return authoredPath(args, "--response", false);
  }

  static Optional<Path> requestPath(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    return authoredPath(args, "--request", true);
  }

  static Optional<Path> responsePathHint(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    try {
      Optional<Path> responsePath = responsePath(args);
      Optional<Path> requestPath = requestPath(args);
      if (responsePath.isPresent()
          && requestPath.isPresent()
          && !isStandardInputPath(requestPath)
          && responsePath
              .orElseThrow()
              .toAbsolutePath()
              .normalize()
              .equals(requestPath.orElseThrow().toAbsolutePath().normalize())) {
        return Optional.empty();
      }
      return responsePath;
    } catch (IllegalArgumentException exception) {
      return Optional.empty();
    }
  }

  static boolean isStandardInputPath(Optional<Path> path) {
    Objects.requireNonNull(path, "path must not be null");
    return path.isPresent() && isStandardInputPath(path.orElseThrow());
  }

  static boolean isStandardInputPath(Path path) {
    Objects.requireNonNull(path, "path must not be null");
    return STANDARD_INPUT_PATH.equals(path.toString());
  }

  static int nextValueIndex(String[] args, int flagIndex, String flagName) {
    int valueIndex = flagIndex + 1;
    if (valueIndex >= args.length) {
      throw new CliArgumentsException(flagName, "Missing value for " + flagName);
    }
    return valueIndex;
  }

  static String requireNonBlankValue(String flagName, String value, String description) {
    if (value.isBlank()) {
      throw new CliArgumentsException(flagName, description + " must not be blank");
    }
    return value;
  }

  static Path requirePathValue(
      String flagName, String value, String description, boolean allowStandardInput) {
    String normalized = requireNonBlankValue(flagName, value, description);
    if (!allowStandardInput && STANDARD_INPUT_PATH.equals(normalized)) {
      throw new CliArgumentsException(
          flagName, flagName + " does not allow the standard-input path sentinel '-'");
    }
    return Path.of(normalized);
  }

  private static Optional<Path> authoredPath(
      String[] args, String flagName, boolean allowStandardInput) {
    Optional<Path> value = Optional.empty();
    int index = 0;
    while (index < args.length) {
      if (!flagName.equals(args[index])) {
        index++;
        continue;
      }
      if (value.isPresent()) {
        throw new CliArgumentsException(flagName, "Duplicate argument: " + flagName);
      }
      int valueIndex = nextValueIndex(args, index, flagName);
      value =
          Optional.of(
              requirePathValue(
                  flagName, args[valueIndex], flagName.substring(2) + " path", allowStandardInput));
      index = valueIndex + 1;
    }
    return value;
  }
}
