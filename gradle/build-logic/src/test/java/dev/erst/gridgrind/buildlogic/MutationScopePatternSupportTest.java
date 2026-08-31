package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies PIT wildcard matching for outer and compiler-generated nested classes. */
class MutationScopePatternSupportTest {
  @Test
  void matchesExactAndWildcardClassPatternsIncludingCompilerGeneratedNestedClasses() {
    List<String> classNames =
        List.of(
            "example.ArchitectureRule",
            "example.ArchitectureRule$1",
            "example.ArchitectureRule$Nested",
            "example.ArchitectureRuleSupport");

    assertEquals(
        List.of(
            "example.ArchitectureRule",
            "example.ArchitectureRule$1",
            "example.ArchitectureRule$Nested",
            "example.ArchitectureRuleSupport"),
        MutationScopePatternSupport.matchingClassNames(classNames, "example.ArchitectureRule*"));
    assertEquals(
        List.of("example.ArchitectureRule"),
        MutationScopePatternSupport.matchingClassNames(classNames, "example.ArchitectureRule"));
    assertTrue(MutationScopePatternSupport.matches("example.ArchitectureRule$1", "example.*$1"));
    assertFalse(MutationScopePatternSupport.matches("example.ArchitectureRule", "example.*$1"));
  }
}
