package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;
import org.junit.jupiter.api.Test;

class JavaSourceShapeReviewedPrefixPolicyTest {
  @Test
  void flagsStalePrefixFamilyHeadroom() {
    JavaSourceShapePolicy.Rule familyRule =
        new JavaSourceShapePolicy.Rule(
            0,
            JavaSourceShapePolicy.MatchKind.PREFIX,
            "executor/src/parityTest/java/dev/erst/gridgrind/engine/runtime/parity/",
            "parity-probe",
            980,
            48,
            null,
            26,
            8,
            8,
            8,
            14,
            "executor",
            JavaSourceShapePolicy.DuplicationGuard.SKIP,
            LocalDate.of(2026, 7, 31),
            "Split parity probes by workbook family before another broad scenario bundle lands.",
            "Parity probe groups should stay split by scenario family.");
    JavaSourceShapePolicy.Rule defaultRule =
        new JavaSourceShapePolicy.Rule(
            1,
            JavaSourceShapePolicy.MatchKind.DEFAULT,
            "*",
            "production-source",
            344,
            22,
            12,
            28,
            13,
            11,
            4,
            12,
            "repo-shape",
            JavaSourceShapePolicy.DuplicationGuard.CHECK,
            null,
            null,
            "Ordinary production Java sources should stay small.");
    SourceShapeMetrics metrics = new SourceShapeMetrics(812, 21, 1, 5, 33, 0, 5, 2, 8);

    List<String> issues =
        JavaSourceShapeReviewPolicy.familyIssues(
            familyRule, metrics, defaultRule, LocalDate.of(2026, 6, 4));

    assertTrue(issues.size() >= 3, () -> "Expected multiple family review issues, but got: " + issues);
    assertTrue(issues.stream().anyMatch(issue -> issue.contains("stale lines headroom 812 -> 980")));
    assertTrue(
        issues.stream().anyMatch(issue -> issue.contains("stale methods headroom 33 -> 48")));
    assertTrue(
        issues.stream()
            .anyMatch(issue -> issue.contains("stale nestedTypes headroom 5 -> 8")));
  }

  @Test
  void flagsExpiredReviewedPrefixFamily() {
    JavaSourceShapePolicy.Rule familyRule =
        new JavaSourceShapePolicy.Rule(
            0,
            JavaSourceShapePolicy.MatchKind.PREFIX,
            "engine/src/main/java/dev/erst/gridgrind/excel/",
            "workbook-support",
            624,
            43,
            27,
            31,
            82,
            39,
            12,
            31,
            "workbook-core",
            JavaSourceShapePolicy.DuplicationGuard.CHECK,
            LocalDate.of(2026, 1, 1),
            "Split workbook support by document concern before broadening the family.",
            "Workbook support remains broad only while the reviewed family stays current.");
    JavaSourceShapePolicy.Rule defaultRule =
        new JavaSourceShapePolicy.Rule(
            1,
            JavaSourceShapePolicy.MatchKind.DEFAULT,
            "*",
            "production-source",
            344,
            22,
            12,
            28,
            13,
            11,
            4,
            12,
            "repo-shape",
            JavaSourceShapePolicy.DuplicationGuard.CHECK,
            null,
            null,
            "Ordinary production Java sources should stay small.");
    SourceShapeMetrics metrics = new SourceShapeMetrics(600, 22, 1, 5, 34, 1, 11, 0, 0);

    List<String> issues =
        JavaSourceShapeReviewPolicy.familyIssues(
            familyRule, metrics, defaultRule, LocalDate.of(2026, 6, 4));

    assertTrue(
        issues.stream().anyMatch(issue -> issue.contains("expired on 2026-01-01")),
        () -> "Expected reviewed family expiry issue, but got: " + issues);
  }
}
