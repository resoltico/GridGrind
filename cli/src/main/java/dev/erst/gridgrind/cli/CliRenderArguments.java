package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Render-oriented CLI argument helpers shared by parsing and fallback reporting. */
final class CliRenderArguments {
  private CliRenderArguments() {}

  static Optional<CliOutputFormat> outputFormat(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    return authoredOutputFormat(args);
  }

  static boolean prettyJsonHint(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    try {
      return authoredPrettyJson(args);
    } catch (IllegalArgumentException exception) {
      return false;
    }
  }

  static GlobalResponseExtraction extractGlobalResponse(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    java.util.List<String> remainingArgs = new java.util.ArrayList<>(args.length);
    Optional<Path> responsePath = Optional.empty();
    Optional<CliOutputFormat> outputFormat = Optional.empty();
    boolean prettyJson = false;
    int index = 0;
    while (index < args.length) {
      String argument = args[index];
      if ("--response".equals(argument)) {
        if (responsePath.isPresent()) {
          throw new CliArgumentsException("--response", "Duplicate argument: --response");
        }
        int valueIndex = CliPathArguments.nextValueIndex(args, index, "--response");
        responsePath =
            Optional.of(
                CliPathArguments.requirePathValue(
                    "--response", args[valueIndex], "response path", false));
        index = valueIndex + 1;
        continue;
      }
      if ("--format".equals(argument)) {
        if (outputFormat.isPresent()) {
          throw new CliArgumentsException("--format", "Duplicate argument: --format");
        }
        int valueIndex = CliPathArguments.nextValueIndex(args, index, "--format");
        String formatValue =
            CliPathArguments.requireNonBlankValue(
                "--format", args[valueIndex], "output format value");
        outputFormat = Optional.of(CliOutputFormat.parse(formatValue));
        index = valueIndex + 1;
        continue;
      }
      if ("--pretty".equals(argument)) {
        if (prettyJson) {
          throw new CliArgumentsException("--pretty", "Duplicate argument: --pretty");
        }
        prettyJson = true;
        index++;
        continue;
      }
      remainingArgs.add(argument);
      index++;
    }
    return new GlobalResponseExtraction(
        List.copyOf(remainingArgs), responsePath, outputFormat, prettyJson);
  }

  private static Optional<CliOutputFormat> authoredOutputFormat(String[] args) {
    Optional<CliOutputFormat> value = Optional.empty();
    int index = 0;
    while (index < args.length) {
      if (!"--format".equals(args[index])) {
        index++;
        continue;
      }
      if (value.isPresent()) {
        throw new CliArgumentsException("--format", "Duplicate argument: --format");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, index, "--format");
      String formatValue =
          CliPathArguments.requireNonBlankValue(
              "--format", args[valueIndex], "output format value");
      value = Optional.of(CliOutputFormat.parse(formatValue));
      index = valueIndex + 1;
    }
    return value;
  }

  private static boolean authoredPrettyJson(String[] args) {
    boolean prettyJson = false;
    int index = 0;
    while (index < args.length) {
      if (!"--pretty".equals(args[index])) {
        index++;
        continue;
      }
      if (prettyJson) {
        throw new CliArgumentsException("--pretty", "Duplicate argument: --pretty");
      }
      prettyJson = true;
      index++;
    }
    return prettyJson;
  }

  record GlobalResponseExtraction(
      List<String> remainingArgs,
      Optional<Path> responsePath,
      Optional<CliOutputFormat> outputFormat,
      boolean prettyJson) {
    GlobalResponseExtraction {
      Objects.requireNonNull(remainingArgs, "remainingArgs must not be null");
      Objects.requireNonNull(responsePath, "responsePath must not be null");
      Objects.requireNonNull(outputFormat, "outputFormat must not be null");
    }

    String[] remainingArgsArray() {
      return remainingArgs.toArray(String[]::new);
    }
  }
}
