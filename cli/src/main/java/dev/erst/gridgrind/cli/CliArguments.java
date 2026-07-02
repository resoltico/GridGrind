package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parses raw command-line arguments into a typed GridGrind CLI command. */
final class CliArguments {
  private CliArguments() {}

  /** Parses the raw CLI args into the corresponding command model. */
  static CliCommand parse(String[] args) {
    return parseInvocation(args).command();
  }

  /** Parses the raw CLI args into one command plus the authored global render options. */
  static CliInvocation parseInvocation(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    CliRenderArguments.GlobalResponseExtraction extraction =
        CliRenderArguments.extractGlobalResponse(args);
    String[] remainingArgs = extraction.remainingArgsArray();
    Optional<Path> responsePath = extraction.responsePath();
    if (remainingArgs.length > 0) {
      Optional<CliImmediateCommandParser.Result> immediate =
          CliImmediateCommandParser.parse(remainingArgs, 0, remainingArgs[0], responsePath);
      if (immediate.isPresent()) {
        CliImmediateCommandParser.Result result = immediate.orElseThrow();
        CliTrailingArgumentValidator.requireNoTrailingArguments(
            remainingArgs, result.nextIndex(), result.commandToken());
        CliInvocation invocation =
            new CliInvocation(result.command(), extraction.outputFormat(), extraction.prettyJson());
        CliRenderOptionValidation.validate(invocation.command(), invocation.outputFormat());
        return invocation;
      }
    }
    CliInvocation invocation =
        new CliInvocation(
            CliExecutionCommandParser.parse(remainingArgs, responsePath),
            extraction.outputFormat(),
            extraction.prettyJson());
    CliRenderOptionValidation.validate(invocation.command(), invocation.outputFormat());
    return invocation;
  }
}
