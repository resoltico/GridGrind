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
  void controlPlaneShapeReviewedSurfacesRemainActiveAndNearTerm() throws IOException {
    JavaSourceShapePolicy policy =
        JavaSourceShapePolicy.load(repositoryPolicy("control-plane-shape-policy.tsv"));
    assertReviewedSurfacesRemainActiveAndNearTerm(
        policy.rules().stream()
            .filter(JavaSourceShapePolicy.Rule::isReviewed)
            .map(JavaSourceShapePolicy.Rule::path)
            .toList(),
        policy.rules().stream()
            .filter(JavaSourceShapePolicy.Rule::isReviewed)
            .map(JavaSourceShapePolicy.Rule::reviewExpiresOn)
            .toList(),
        "control-plane",
        "Control-plane shape");
  }

  @Test
  void sourceShapeReviewedSurfacesRemainActiveAndNearTerm() throws IOException {
    JavaSourceShapePolicy policy = JavaSourceShapePolicy.load(repositoryPolicy("source-shape-policy.tsv"));
    assertReviewedSurfacesRemainActiveAndNearTerm(
        policy.rules().stream()
            .filter(JavaSourceShapePolicy.Rule::isReviewed)
            .map(JavaSourceShapePolicy.Rule::path)
            .toList(),
        policy.rules().stream()
            .filter(JavaSourceShapePolicy.Rule::isReviewed)
            .map(JavaSourceShapePolicy.Rule::reviewExpiresOn)
            .toList(),
        "repository",
        "Source-shape");
  }

  @Test
  void semanticShapeReviewedSurfacesRemainActiveAndNearTerm() throws IOException {
    JavaSemanticShapePolicy policy =
        JavaSemanticShapePolicy.load(repositoryPolicy("semantic-shape-policy.tsv"));
    assertReviewedSurfacesRemainActiveAndNearTerm(
        policy.rules().stream()
            .map(JavaSemanticShapePolicy.Rule::path)
            .toList(),
        policy.rules().stream()
            .map(JavaSemanticShapePolicy.Rule::reviewExpiresOn)
            .toList(),
        "semantic",
        "Semantic-shape");
  }

  private static void assertReviewedSurfacesRemainActiveAndNearTerm(
      List<String> reviewedPaths,
      List<LocalDate> reviewDates,
      String debtLabel,
      String policyName) {
    LocalDate today = LocalDate.now(ZoneOffset.UTC);
    LocalDate maxAllowed = today.plusDays(MAX_REVIEW_HORIZON_DAYS);
    List<String> issues = new ArrayList<>();
    for (int index = 0; index < reviewDates.size(); index++) {
      LocalDate reviewDate = reviewDates.get(index);
      String reviewedPath = reviewedPaths.get(index);
      if (reviewDate.isBefore(today)) {
        issues.add(reviewedPath + " expired on " + reviewDate);
      }
      if (reviewDate.isAfter(maxAllowed)) {
        issues.add(
            reviewedPath
                + " parks "
                + debtLabel
                + " review debt too far out at "
                + reviewDate
                + " (max "
                + maxAllowed
                + ").");
      }
    }
    assertFalse(
        reviewedPaths.isEmpty(),
        policyName + " policy defines no reviewed surfaces; add at least one live reviewed surface or remove this guard.");
    assertTrue(issues.isEmpty(), () -> policyName + " policy freshness issues: " + issues);
  }

  private static Path repositoryPolicy(String fileName) {
    String repositoryRoot = System.getProperty("gridgrind.repository.root");
    if (repositoryRoot == null || repositoryRoot.isBlank()) {
      throw new IllegalStateException("gridgrind.repository.root system property must be set for build-logic tests.");
    }
    return Path.of(repositoryRoot).resolve("gradle").resolve(fileName);
  }
}
