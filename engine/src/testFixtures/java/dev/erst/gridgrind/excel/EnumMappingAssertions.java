package dev.erst.gridgrind.excel;

import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Function;

/** Shared assertion helpers for enum and bridge mapping coverage. */
final class EnumMappingAssertions {
  private EnumMappingAssertions() {}

  static <Source, Target> void assertMappings(
      Map<Source, Target> expectedMappings, Function<Source, Target> mapper) {
    expectedMappings.forEach(
        (source, expectedTarget) -> assertSameValue(expectedTarget, mapper.apply(source)));
  }

  static <Source, Target> void assertBidirectionalMappings(
      Map<Source, Target> expectedMappings,
      Function<Source, Target> fromSource,
      Function<Target, Source> toSource) {
    expectedMappings.forEach(
        (source, expectedTarget) -> {
          assertSameValue(expectedTarget, fromSource.apply(source));
          assertSameValue(source, toSource.apply(expectedTarget));
        });
  }

  static <Source, Target> void assertOptionalBidirectionalMappings(
      Map<Source, Target> expectedMappings,
      Function<Source, Optional<Target>> fromSource,
      Function<Target, Source> toSource) {
    expectedMappings.forEach(
        (source, expectedTarget) -> {
          assertSameValue(Optional.of(expectedTarget), fromSource.apply(source));
          assertSameValue(source, toSource.apply(expectedTarget));
        });
  }

  private static void assertSameValue(Object expected, Object actual) {
    if (!Objects.equals(expected, actual)) {
      throw new AssertionError("Expected " + expected + " but was " + actual);
    }
  }
}
