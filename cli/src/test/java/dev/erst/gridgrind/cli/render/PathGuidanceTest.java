package dev.erst.gridgrind.cli.render;

import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Verifies that host guidance preserves explicit request-root and stream ownership. */
class PathGuidanceTest {
  @Test
  void namesTheHostOptionsThatOwnStdinAndRelativePaths() {
    assertTrue(PathGuidance.stdinExecutionRootRequiredMessage().contains("--execution-root"));
    assertTrue(PathGuidance.cliFlagPathResolutionSummary().contains("current working directory"));
    assertTrue(PathGuidance.standardInputRequiresRequestMessage().contains("STANDARD_INPUT"));
    assertTrue(PathGuidance.standardInputRequiresRequestMessage().contains("--request"));
  }
}
