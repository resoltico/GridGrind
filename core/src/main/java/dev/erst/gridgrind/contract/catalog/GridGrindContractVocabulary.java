package dev.erst.gridgrind.contract.catalog;

import java.util.Set;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/** Canonical registered public operation, assertion, and inspection ids. */
public final class GridGrindContractVocabulary {
  private static final Pattern OPERATION_ID_PATTERN = Pattern.compile("[A-Z]+(?:_[A-Z0-9*]*)+");

  private GridGrindContractVocabulary() {}

  /**
   * Returns every canonical mutation, assertion, and inspection id that public surfaces may use.
   */
  public static Set<String> registeredPublicSurfaceIds() {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    return Stream.of(
            catalog.mutationActionTypes().stream(),
            catalog.assertionTypes().stream(),
            catalog.inspectionQueryTypes().stream())
        .flatMap(stream -> stream)
        .map(TypeEntry::id)
        .collect(Collectors.toUnmodifiableSet());
  }

  /** Returns a pattern that matches SCREAMING_SNAKE_CASE candidate operation ids. */
  public static Pattern candidateIdPattern() {
    return OPERATION_ID_PATTERN;
  }
}
