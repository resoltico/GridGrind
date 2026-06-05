package dev.erst.gridgrind.buildlogic;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

final class JavaSourceShapeReviewPolicy {
  private static final int MAX_LINE_HEADROOM = 24;
  private static final int MAX_METHOD_HEADROOM = 3;
  private static final int MAX_PUBLIC_METHOD_HEADROOM = 3;
  private static final int MAX_IMPORT_HEADROOM = 4;
  private static final int MAX_FIELD_HEADROOM = 2;
  private static final int MAX_SWITCH_HEADROOM = 2;
  private static final int MAX_SWITCH_ARM_HEADROOM = 8;

  private JavaSourceShapeReviewPolicy() {}

  static List<String> policyIssues(
      String relativePath,
      JavaSourceShapePolicy.Rule reviewedRule,
      JavaSourceShapePolicy.Rule broaderRule,
      JavaSourceShapeAnalyzer.Metrics metrics,
      LocalDate today) {
    List<String> issues = new ArrayList<>();
    if (reviewedRule.reviewExpiresOn().isBefore(today)) {
      issues.add(
          "EXACT reviewed surface for "
              + relativePath
              + " expired on "
              + reviewedRule.reviewExpiresOn()
              + ". Re-review the surface or drop the exact override.");
    }
    if (broaderRule != null && !meaningfullyStricterThan(reviewedRule, broaderRule)) {
      issues.add(
          "EXACT reviewed surface for "
              + relativePath
              + " does not tighten any budget beyond broader rule "
              + broaderRule.role()
              + ". Tighten at least one metric or remove the exact override.");
    }
    addHeadroomIssue(
        issues,
        "lines",
        relativePath,
        metrics.lineCount(),
        reviewedRule.maxLines(),
        MAX_LINE_HEADROOM);
    addHeadroomIssue(
        issues,
        "methods",
        relativePath,
        metrics.methodCount(),
        reviewedRule.maxMethods(),
        MAX_METHOD_HEADROOM);
    addHeadroomIssue(
        issues,
        "publicMethods",
        relativePath,
        metrics.publicMethodCount(),
        reviewedRule.maxPublicMethods(),
        MAX_PUBLIC_METHOD_HEADROOM);
    addHeadroomIssue(
        issues,
        "imports",
        relativePath,
        metrics.importCount(),
        reviewedRule.maxImports(),
        MAX_IMPORT_HEADROOM);
    addHeadroomIssue(
        issues,
        "fields",
        relativePath,
        metrics.fieldCount(),
        reviewedRule.maxFields(),
        MAX_FIELD_HEADROOM);
    addHeadroomIssue(
        issues,
        "switches",
        relativePath,
        metrics.switchCount(),
        reviewedRule.maxSwitches(),
        MAX_SWITCH_HEADROOM);
    addHeadroomIssue(
        issues,
        "maxSwitchArms",
        relativePath,
        metrics.maxSwitchArms(),
        reviewedRule.maxSwitchArms(),
        MAX_SWITCH_ARM_HEADROOM);
    return List.copyOf(issues);
  }

  private static boolean meaningfullyStricterThan(
      JavaSourceShapePolicy.Rule reviewedRule, JavaSourceShapePolicy.Rule broaderRule) {
    return tighter(reviewedRule.maxLines(), broaderRule.maxLines())
        || tighter(reviewedRule.maxMethods(), broaderRule.maxMethods())
        || tighter(reviewedRule.maxPublicMethods(), broaderRule.maxPublicMethods())
        || tighter(reviewedRule.maxImports(), broaderRule.maxImports())
        || tighter(reviewedRule.maxFields(), broaderRule.maxFields())
        || tighter(reviewedRule.maxSwitches(), broaderRule.maxSwitches())
        || tighter(reviewedRule.maxSwitchArms(), broaderRule.maxSwitchArms());
  }

  private static boolean tighter(Integer reviewedLimit, Integer broaderLimit) {
    return reviewedLimit != null && (broaderLimit == null || reviewedLimit < broaderLimit);
  }

  private static void addHeadroomIssue(
      List<String> issues,
      String metricName,
      String relativePath,
      long currentValue,
      Integer reviewedLimit,
      int allowedHeadroom) {
    if (reviewedLimit == null) {
      return;
    }
    long headroom = (long) reviewedLimit - currentValue;
    if (headroom > allowedHeadroom) {
      issues.add(
          "EXACT reviewed surface for "
              + relativePath
              + " carries stale "
              + metricName
              + " headroom "
              + currentValue
              + " -> "
              + reviewedLimit
              + " (allowed "
              + allowedHeadroom
              + "). Tighten the reviewed budget.");
    }
  }
}
