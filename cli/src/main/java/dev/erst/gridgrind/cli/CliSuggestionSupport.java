package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Derives runnable CLI suggestions from owned contract facts instead of canned examples. */
final class CliSuggestionSupport {
  private static final Pattern FIRST_QUOTED_TOKEN = Pattern.compile("'([^']+)'");
  private static final Pattern ARRAY_INDEX = Pattern.compile("\\[\\d+\\]");

  private CliSuggestionSupport() {}

  static Optional<String> protocolCatalogSearchCommandForLookupId(String lookupId) {
    Objects.requireNonNull(lookupId, "lookupId must not be null");
    return nonBlank(lookupId).map(CliSuggestionSupport::protocolCatalogSearchCommand);
  }

  static Optional<String> protocolCatalogSearchCommandForProblem(
      GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    Optional<String> contextualQuery = Optional.empty();
    if (problem.context() instanceof ProblemContext.ReadRequest readRequest) {
      contextualQuery =
          searchQueryForJsonPath(readRequest.jsonPath())
              .or(() -> searchQueryFromMessage(problem.message()));
    }
    return contextualQuery.map(CliSuggestionSupport::protocolCatalogSearchCommand);
  }

  private static Optional<String> searchQueryForJsonPath(Optional<String> jsonPath) {
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    return jsonPath
        .map(CliSuggestionSupport::normalizeJsonPathQuery)
        .flatMap(CliSuggestionSupport::nonBlank);
  }

  private static Optional<String> searchQueryFromMessage(String message) {
    Objects.requireNonNull(message, "message must not be null");
    Matcher matcher = FIRST_QUOTED_TOKEN.matcher(message);
    if (!matcher.find()) {
      return Optional.empty();
    }
    return nonBlank(matcher.group(1));
  }

  private static String normalizeJsonPathQuery(String jsonPath) {
    String normalized =
        ARRAY_INDEX.matcher(jsonPath).replaceAll(" ").replace('.', ' ').replace('_', ' ');
    normalized = normalized.replaceAll("\\s+", " ").trim();
    return normalized.endsWith(" type")
        ? normalized.substring(0, normalized.length() - " type".length()).trim() + " type"
        : normalized;
  }

  private static Optional<String> nonBlank(String value) {
    String normalized = value.trim();
    return normalized.isBlank() ? Optional.empty() : Optional.of(normalized);
  }

  static String protocolCatalogSearchCommand(String query) {
    return "gridgrind --print-protocol-catalog --search \"" + escapeQuoted(query) + "\"";
  }

  private static String escapeQuoted(String query) {
    return query.replace("\\", "\\\\").replace("\"", "\\\"");
  }
}
