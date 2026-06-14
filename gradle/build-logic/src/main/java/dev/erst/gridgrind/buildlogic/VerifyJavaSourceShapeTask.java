package dev.erst.gridgrind.buildlogic;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Set;
import org.gradle.api.DefaultTask;
import org.gradle.api.GradleException;
import org.gradle.api.file.ConfigurableFileCollection;
import org.gradle.api.provider.Property;
import org.gradle.api.file.RegularFileProperty;
import org.gradle.api.tasks.Input;
import org.gradle.api.tasks.InputFile;
import org.gradle.api.tasks.InputFiles;
import org.gradle.api.tasks.OutputFile;
import org.gradle.api.tasks.PathSensitive;
import org.gradle.api.tasks.PathSensitivity;
import org.gradle.api.tasks.TaskAction;

public abstract class VerifyJavaSourceShapeTask extends DefaultTask {
  @InputFiles
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract ConfigurableFileCollection getSourceRoots();

  @Input
  public abstract Property<Integer> getJavaRelease();

  @Input
  public abstract Property<String> getRepositoryRootPath();

  @Input
  public abstract Property<String> getReviewDate();

  @InputFile
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract RegularFileProperty getPolicyFile();

  @OutputFile
  public abstract RegularFileProperty getReportFile();

  @TaskAction
  void verifySourceShape() throws IOException {
    Path repositoryRoot = Path.of(getRepositoryRootPath().get()).toAbsolutePath().normalize();
    Path policyPath = getPolicyFile().get().getAsFile().toPath().toAbsolutePath().normalize();
    Path reportPath = getReportFile().get().getAsFile().toPath().toAbsolutePath().normalize();
    int javaRelease = getJavaRelease().get();
    JavaSourceShapePolicy policy = JavaSourceShapePolicy.load(policyPath);
    JavaSourceShapePolicy.Rule defaultRule = policy.rules().getLast();
    JavaSourceShapeAnalyzer analyzer = new JavaSourceShapeAnalyzer();
    LocalDate today = LocalDate.parse(getReviewDate().get());

    List<Path> sourceFiles = collectSourceFiles();
    List<String> relativePaths =
        sourceFiles.stream().map(repositoryRoot::relativize).map(this::normalizePath).toList();
    List<String> policyIssues = new ArrayList<>();
    List<Violation> violations = new ArrayList<>();
    List<ReportRow> reportRows = new ArrayList<>();
    java.util.Map<JavaSourceShapePolicy.Rule, FamilyMetrics> familyMetrics =
        new java.util.LinkedHashMap<>();

    for (JavaSourceShapePolicy.Rule rule : policy.rules()) {
      if (rule.kind() == JavaSourceShapePolicy.MatchKind.EXACT) {
        if (!Files.isRegularFile(repositoryRoot.resolve(rule.path()))) {
          policyIssues.add("EXACT rule points at a missing file: " + rule.path());
        }
      }
      if (rule.kind() == JavaSourceShapePolicy.MatchKind.PREFIX
          && relativePaths.stream().noneMatch(rule::matches)) {
        policyIssues.add("PREFIX rule matches no repo-owned Java source files: " + rule.path());
      }
    }

    for (int index = 0; index < sourceFiles.size(); index++) {
      Path sourceFile = sourceFiles.get(index);
      String relativePath = relativePaths.get(index);
      List<JavaSourceShapePolicy.Rule> matchingRules = policy.matchingRules(relativePath);
      if (matchingRules.isEmpty()) {
        policyIssues.add("No source-shape rule matches " + relativePath);
        continue;
      }

      JavaSourceShapePolicy.Rule matchedRule = matchingRules.getFirst();
      SourceShapeMetrics metrics = analyzer.analyze(sourceFile, javaRelease);
      matchingRules.stream()
          .filter(rule -> rule.kind() == JavaSourceShapePolicy.MatchKind.PREFIX)
          .forEach(
              rule ->
                  familyMetrics.computeIfAbsent(rule, unused -> new FamilyMetrics()).include(metrics));
      List<String> exceededMetrics = SourceShapeBudgetSupport.exceededMetrics(metrics, matchedRule);
      if (!exceededMetrics.isEmpty()) {
        violations.add(new Violation(relativePath, matchedRule, metrics, exceededMetrics));
      }

      if (matchedRule.kind() == JavaSourceShapePolicy.MatchKind.EXACT) {
        JavaSourceShapePolicy.Rule broaderRule =
            matchingRules.stream()
                .skip(1)
                .filter(rule -> rule.kind() != JavaSourceShapePolicy.MatchKind.EXACT)
                .findFirst()
                .orElse(null);
        policyIssues.addAll(
            JavaSourceShapeReviewPolicy.policyIssues(
                relativePath, matchedRule, broaderRule, metrics, today));
      }

      reportRows.add(new ReportRow(relativePath, matchedRule, metrics));
    }

    for (JavaSourceShapePolicy.Rule rule : policy.rules()) {
      if (rule.kind() != JavaSourceShapePolicy.MatchKind.PREFIX) {
        continue;
      }
      FamilyMetrics metrics = familyMetrics.get(rule);
      if (metrics == null) {
        continue;
      }
      policyIssues.addAll(
          JavaSourceShapeReviewPolicy.familyIssues(
              rule, metrics.toMetrics(), defaultRule, today));
    }

    SourceShapeBudgetSupport.writeBudgetReport(reportPath, reportRows);

    if (!policyIssues.isEmpty() || !violations.isEmpty()) {
      throw new GradleException(
          buildFailureMessage(reportPath, policyIssues, violations));
    }
  }

  private List<Path> collectSourceFiles() throws IOException {
    List<Path> sourceFiles = new ArrayList<>();
    Set<File> roots = getSourceRoots().getFiles();
    for (File sourceRoot : roots) {
      if (!sourceRoot.isDirectory()) {
        continue;
      }
      try (var walked = Files.walk(sourceRoot.toPath())) {
        walked
            .filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".java"))
            .sorted()
            .forEach(sourceFiles::add);
      }
    }
    sourceFiles.sort(Comparator.naturalOrder());
    return sourceFiles;
  }

  private String buildFailureMessage(
      Path reportPath, List<String> policyIssues, List<Violation> violations) {
    StringBuilder message = new StringBuilder();
    message
        .append("Java source-shape verification failed.")
        .append(System.lineSeparator())
        .append("Report: ")
        .append(reportPath)
        .append(System.lineSeparator());
    if (!policyIssues.isEmpty()) {
      message.append("Policy issues:").append(System.lineSeparator());
      for (String issue : policyIssues) {
        message.append(" - ").append(issue).append(System.lineSeparator());
      }
    }
    if (!violations.isEmpty()) {
      message.append("Budget violations:").append(System.lineSeparator());
      for (Violation violation : violations.stream().limit(25).toList()) {
        message
            .append(" - ")
            .append(violation.relativePath())
            .append(" [")
            .append(violation.rule().role())
            .append("] exceeded ")
            .append(String.join(", ", violation.exceededMetrics()))
            .append(System.lineSeparator());
      }
      if (violations.size() > 25) {
        message
            .append(" - ... and ")
            .append(violations.size() - 25)
            .append(" more. See the TSV report for the full list.")
            .append(System.lineSeparator());
      }
    }
    return message.toString();
  }

  private String normalizePath(Path relativePath) {
    return relativePath.toString().replace(File.separatorChar, '/');
  }

  private record Violation(
      String relativePath,
      JavaSourceShapePolicy.Rule rule,
      SourceShapeMetrics metrics,
      List<String> exceededMetrics) {}

  private record ReportRow(
      String relativePath,
      JavaSourceShapePolicy.Rule rule,
      SourceShapeMetrics metrics)
      implements SourceShapeBudgetSupport.ReportRowView {
    @Override
    public double riskScore() {
      return SourceShapeBudgetSupport.riskScore(metrics, rule);
    }

    @Override
    public String toTsv() {
      return SourceShapeBudgetSupport.reportRowTsv(relativePath, rule, metrics);
    }
  }

  private static final class FamilyMetrics {
    private long lineCount;
    private int importCount;
    private int topLevelTypeCount;
    private int nestedTypeCount;
    private int methodCount;
    private int publicMethodCount;
    private int fieldCount;
    private int switchCount;
    private int maxSwitchArms;

    private void include(SourceShapeMetrics metrics) {
      lineCount = Math.max(lineCount, metrics.lineCount());
      importCount = Math.max(importCount, metrics.importCount());
      topLevelTypeCount = Math.max(topLevelTypeCount, metrics.topLevelTypeCount());
      nestedTypeCount = Math.max(nestedTypeCount, metrics.nestedTypeCount());
      methodCount = Math.max(methodCount, metrics.methodCount());
      publicMethodCount = Math.max(publicMethodCount, metrics.publicMethodCount());
      fieldCount = Math.max(fieldCount, metrics.fieldCount());
      switchCount = Math.max(switchCount, metrics.switchCount());
      maxSwitchArms = Math.max(maxSwitchArms, metrics.maxSwitchArms());
    }

    private SourceShapeMetrics toMetrics() {
      return new SourceShapeMetrics(
          lineCount,
          importCount,
          topLevelTypeCount,
          nestedTypeCount,
          methodCount,
          publicMethodCount,
          fieldCount,
          switchCount,
          maxSwitchArms);
    }
  }
}
