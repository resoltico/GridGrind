package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Centralizes parse-argument failure suggestions and operator recovery guidance. */
final class CliArgumentFailureSupport {
  private static final List<String> KNOWN_FLAGS =
      List.of(
          "--request",
          "--response",
          "--doctor-request",
          "--print-request-template",
          "--print-example-catalog",
          "--print-example",
          "--print-task-catalog",
          "--print-task-plan",
          "--print-task-keyword-match",
          "--print-protocol-catalog",
          "--lookup",
          "--query",
          "--search",
          "--help",
          "--help-protocol",
          "--help-guidance",
          "--version",
          "--license");

  private CliArgumentFailureSupport() {}

  static CliFailureReport reportFor(CliArgumentsException exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    String message =
        Objects.requireNonNullElse(exception.getMessage(), "Invalid command-line arguments");
    return CliFailureReports.invalidArguments(
        2,
        "parse-arguments",
        Optional.of(exception.argument()),
        message,
        suggestionsFor(Optional.of(exception.argument()), message),
        resolutionFor(Optional.of(exception.argument()), message));
  }

  static CliFailureReport reportFor(IllegalArgumentException exception) {
    Objects.requireNonNull(exception, "exception must not be null");
    String message =
        Objects.requireNonNullElse(exception.getMessage(), "Invalid command-line arguments");
    return CliFailureReports.invalidArguments(
        2,
        "parse-arguments",
        Optional.empty(),
        message,
        suggestionsFor(Optional.empty(), message),
        resolutionFor(Optional.empty(), message));
  }

  private static List<String> suggestionsFor(Optional<String> argument, String message) {
    Set<String> suggestions = new LinkedHashSet<>();
    argument.ifPresent(value -> suggestions.addAll(argumentSuggestions(value, message)));
    if (suggestions.isEmpty()) {
      suggestions.add("gridgrind --help");
      suggestions.add("gridgrind --help-protocol");
      suggestions.add("gridgrind --help-guidance");
    }
    return List.copyOf(suggestions);
  }

  private static List<String> argumentSuggestions(String argument, String message) {
    Set<String> suggestions = new LinkedHashSet<>();
    switch (argument) {
      case "--request" -> {
        suggestions.add("gridgrind --request request.json --response response.json");
        suggestions.add("gridgrind --doctor-request --request request.json --response doctor.json");
      }
      case "--response" -> {
        suggestions.add("gridgrind --response response.json");
        suggestions.add("gridgrind --help-guidance");
      }
      case "--lookup" -> {
        suggestions.add("gridgrind --print-example-catalog");
        suggestions.add("gridgrind --print-task-catalog");
        suggestions.add("gridgrind --print-protocol-catalog --search <text>");
      }
      case "--query" -> {
        suggestions.add("gridgrind --print-task-keyword-match --query \"monthly sales dashboard\"");
        suggestions.add("gridgrind --print-task-catalog");
      }
      case "--search" -> {
        suggestions.add("gridgrind --print-protocol-catalog --search <text>");
        suggestions.add("gridgrind --help-protocol");
      }
      default -> {
        if (message.startsWith("Unknown argument: ")) {
          nearestFlags(argument).stream()
              .map(CliArgumentFailureSupport::commandTemplatesForFlag)
              .flatMap(List::stream)
              .forEach(suggestions::add);
        }
      }
    }
    if (suggestions.isEmpty()) {
      suggestions.add("gridgrind --help");
      suggestions.add("gridgrind --help-protocol");
      suggestions.add("gridgrind --help-guidance");
    }
    return List.copyOf(suggestions);
  }

  private static Optional<String> resolutionFor(Optional<String> argument, String message) {
    Objects.requireNonNull(argument, "argument must not be null");
    if (message.startsWith("Unknown argument: ")) {
      return Optional.of(
          "Use one exact CLI flag. Start from --help for the synopsis, --help-protocol for the"
              + " grammar, or --help-guidance for workflow-oriented commands.");
    }
    return argument
        .map(
            value ->
                switch (value) {
                  case "--request" ->
                      "Provide one readable request JSON file path, or omit --request and pipe one"
                          + " request document on standard input.";
                  case "--response" -> "Provide one writable response file path after --response.";
                  case "--lookup" ->
                      "Use --lookup only with --print-example, --print-task-catalog,"
                          + " --print-task-plan, or --print-protocol-catalog.";
                  case "--query" ->
                      "Use --query only with --print-task-keyword-match and provide one natural-language query.";
                  case "--search" ->
                      "Use --search only with --print-protocol-catalog and provide one search string.";
                  default ->
                      "Run gridgrind --help for the synopsis, --help-protocol for the authoritative"
                          + " request contract, or --help-guidance for workflows and examples.";
                })
        .or(
            () ->
                Optional.of(
                    "Run gridgrind --help for the synopsis, --help-protocol for the authoritative"
                        + " request contract, or --help-guidance for workflows and examples."));
  }

  private static List<String> nearestFlags(String argument) {
    String normalized = normalize(argument);
    return KNOWN_FLAGS.stream()
        .sorted(
            java.util.Comparator.comparingInt(
                    (String flag) -> distance(normalized, normalize(flag)))
                .thenComparing(String::length)
                .thenComparing(String::compareTo))
        .filter(flag -> distance(normalized, normalize(flag)) <= 2)
        .limit(3)
        .toList();
  }

  static List<String> commandTemplatesForFlag(String flag) {
    return switch (flag) {
      case "--request" -> List.of("gridgrind --request request.json --response response.json");
      case "--response" -> List.of("gridgrind --print-request-template --response request.json");
      case "--doctor-request" ->
          List.of("gridgrind --doctor-request --request request.json --response doctor.json");
      case "--print-request-template" ->
          List.of("gridgrind --print-request-template --response request.json");
      case "--print-example-catalog" -> List.of("gridgrind --print-example-catalog");
      case "--print-example" -> List.of("gridgrind --print-example --lookup WORKBOOK_HEALTH");
      case "--print-task-catalog" -> List.of("gridgrind --print-task-catalog");
      case "--print-task-plan" -> List.of("gridgrind --print-task-plan --lookup DASHBOARD");
      case "--print-task-keyword-match" ->
          List.of("gridgrind --print-task-keyword-match --query \"monthly sales dashboard\"");
      case "--print-protocol-catalog" ->
          List.of("gridgrind --print-protocol-catalog --search \"sheet layout\"");
      case "--lookup" ->
          List.of(
              "gridgrind --print-example --lookup WORKBOOK_HEALTH",
              "gridgrind --print-task-plan --lookup DASHBOARD",
              "gridgrind --print-protocol-catalog --lookup GET_CELLS");
      case "--query" ->
          List.of("gridgrind --print-task-keyword-match --query \"monthly sales dashboard\"");
      case "--search" -> List.of("gridgrind --print-protocol-catalog --search \"sheet layout\"");
      case "--help", "--help-protocol", "--help-guidance", "--version", "--license" ->
          List.of("gridgrind " + flag);
      default -> List.of();
    };
  }

  private static String normalize(String value) {
    return value.replace("-", "").toLowerCase(Locale.ROOT);
  }

  private static int distance(String left, String right) {
    int[][] matrix = new int[left.length() + 1][right.length() + 1];
    for (int i = 0; i <= left.length(); i++) {
      matrix[i][0] = i;
    }
    for (int j = 0; j <= right.length(); j++) {
      matrix[0][j] = j;
    }
    for (int i = 1; i <= left.length(); i++) {
      for (int j = 1; j <= right.length(); j++) {
        int substitutionCost = left.charAt(i - 1) == right.charAt(j - 1) ? 0 : 1;
        matrix[i][j] =
            Math.min(
                Math.min(matrix[i - 1][j] + 1, matrix[i][j - 1] + 1),
                matrix[i - 1][j - 1] + substitutionCost);
      }
    }
    return matrix[left.length()][right.length()];
  }
}
