package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;
import org.junit.jupiter.api.Test;

class JavaSourceShapeReviewPolicyTest {
  @Test
  void flagsExpiredAndOverbroadReviewedSurface() {
    JavaSourceShapePolicy.Rule reviewedRule =
        new JavaSourceShapePolicy.Rule(
            0,
            JavaSourceShapePolicy.MatchKind.EXACT,
            "cli/src/main/java/dev/erst/gridgrind/cli/GridGrindTaskKeywordMatcher.java",
            "cli-keyword-match",
            520,
            30,
            18,
            30,
            null,
            10,
            16,
            "cli",
            JavaSourceShapePolicy.DuplicationGuard.CHECK,
            LocalDate.of(2026, 1, 1),
            "Split normalization from scoring before adding another discovery surface.",
            "Keyword matching remains one reviewed discovery surface.");
    JavaSourceShapePolicy.Rule broaderRule =
        new JavaSourceShapePolicy.Rule(
            1,
            JavaSourceShapePolicy.MatchKind.PREFIX,
            "cli/src/main/java/dev/erst/gridgrind/cli/",
            "cli-support",
            500,
            28,
            18,
            30,
            null,
            10,
            16,
            "cli",
            JavaSourceShapePolicy.DuplicationGuard.CHECK,
            null,
            null,
            "CLI orchestration aggregates command wiring and help surfaces.");
    JavaSourceShapeAnalyzer.Metrics metrics =
        new JavaSourceShapeAnalyzer.Metrics(492, 27, 1, 0, 26, 0, 8, 9, 16);

    List<String> issues =
        JavaSourceShapeReviewPolicy.policyIssues(
            reviewedRule.path(), reviewedRule, broaderRule, metrics, LocalDate.of(2026, 6, 4));

    assertTrue(issues.size() >= 3, () -> "Expected multiple review issues, but got: " + issues);
    assertTrue(issues.get(0).contains("expired on 2026-01-01"));
    assertTrue(issues.stream().anyMatch(issue -> issue.contains("does not tighten any budget")));
    assertTrue(issues.stream().anyMatch(issue -> issue.contains("stale lines headroom 492 -> 520")));
  }

  @Test
  void acceptsTightReviewedSurfaceThatAddsRealConstraint() {
    JavaSourceShapePolicy.Rule reviewedRule =
        new JavaSourceShapePolicy.Rule(
            0,
            JavaSourceShapePolicy.MatchKind.EXACT,
            "contract/src/main/java/dev/erst/gridgrind/contract/json/GridGrindJson.java",
            "contract-json-codec",
            256,
            32,
            24,
            16,
            null,
            2,
            4,
            "contract",
            JavaSourceShapePolicy.DuplicationGuard.CHECK,
            LocalDate.of(2026, 9, 30),
            "Split request, response, and discovery codec families before adding another public wire payload family.",
            "GridGrindJson remains the singular protocol JSON facade.");
    JavaSourceShapePolicy.Rule broaderRule =
        new JavaSourceShapePolicy.Rule(
            1,
            JavaSourceShapePolicy.MatchKind.PREFIX,
            "contract/src/main/java/dev/erst/gridgrind/contract/json/",
            "contract-json",
            320,
            32,
            24,
            20,
            null,
            8,
            12,
            "contract",
            JavaSourceShapePolicy.DuplicationGuard.CHECK,
            null,
            null,
            "JSON translation surfaces own the published wire contract.");
    JavaSourceShapeAnalyzer.Metrics metrics =
        new JavaSourceShapeAnalyzer.Metrics(251, 14, 1, 0, 32, 24, 0, 0, 0);

    List<String> issues =
        JavaSourceShapeReviewPolicy.policyIssues(
            reviewedRule.path(), reviewedRule, broaderRule, metrics, LocalDate.of(2026, 6, 4));

    assertTrue(issues.isEmpty(), () -> "Expected no review issues, but got: " + issues);
  }
}
