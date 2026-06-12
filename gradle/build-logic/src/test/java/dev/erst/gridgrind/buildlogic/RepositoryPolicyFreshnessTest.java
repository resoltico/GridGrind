package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertFalse;

import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.ZoneOffset;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RepositoryPolicyFreshnessTest {
  private static final int MAX_REVIEW_HORIZON_DAYS = 180;

  @Test
  void sourceShapeReviewedSurfacesRemainActiveAndNearTerm() throws IOException {
    JavaSourceShapePolicy policy = JavaSourceShapePolicy.load(repositoryPolicy("source-shape-policy.tsv"));
    LocalDate today = LocalDate.now(ZoneOffset.UTC);
    LocalDate maxAllowed = today.plusDays(MAX_REVIEW_HORIZON_DAYS);

    List<String> issues = new ArrayList<>();
    int reviewedCount = 0;
    for (JavaSourceShapePolicy.Rule rule : policy.rules()) {
      if (rule.kind() != JavaSourceShapePolicy.MatchKind.EXACT) {
        continue;
      }
      reviewedCount++;
      if (rule.reviewExpiresOn().isBefore(today)) {
        issues.add(rule.path() + " expired on " + rule.reviewExpiresOn());
      }
      if (rule.reviewExpiresOn().isAfter(maxAllowed)) {
        issues.add(
            rule.path()
                + " parks review debt too far out at "
                + rule.reviewExpiresOn()
                + " (max "
                + maxAllowed
                + ").");
      }
    }

    assertTrue(
        reviewedCount > 0,
        "Repository source-shape policy defines no reviewed exact surfaces; add at least one live reviewed surface or remove this guard.");
    assertTrue(issues.isEmpty(), () -> "Source-shape policy freshness issues: " + issues);
  }

  @Test
  void semanticShapeReviewedSurfacesRemainActiveAndNearTerm() throws IOException {
    JavaSemanticShapePolicy policy =
        JavaSemanticShapePolicy.load(repositoryPolicy("semantic-shape-policy.tsv"));
    LocalDate today = LocalDate.now(ZoneOffset.UTC);
    LocalDate maxAllowed = today.plusDays(MAX_REVIEW_HORIZON_DAYS);

    List<String> issues = new ArrayList<>();
    int reviewedCount = 0;
    for (JavaSemanticShapePolicy.Rule rule : policy.rules()) {
      reviewedCount++;
      if (rule.reviewExpiresOn().isBefore(today)) {
        issues.add(rule.path() + " expired on " + rule.reviewExpiresOn());
      }
      if (rule.reviewExpiresOn().isAfter(maxAllowed)) {
        issues.add(
            rule.path()
                + " parks semantic review debt too far out at "
                + rule.reviewExpiresOn()
                + " (max "
                + maxAllowed
                + ").");
      }
    }

    assertTrue(
        reviewedCount > 0,
        "Repository semantic-shape policy defines no reviewed surfaces; add at least one live reviewed surface or remove this guard.");
    assertTrue(issues.isEmpty(), () -> "Semantic-shape policy freshness issues: " + issues);
  }

  private static Path repositoryPolicy(String fileName) {
    String repositoryRoot = System.getProperty("gridgrind.repository.root");
    if (repositoryRoot == null || repositoryRoot.isBlank()) {
      throw new IllegalStateException("gridgrind.repository.root system property must be set for build-logic tests.");
    }
    return Path.of(repositoryRoot).resolve("gradle").resolve(fileName);
  }
}
