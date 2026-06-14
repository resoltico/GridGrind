package dev.erst.gridgrind.buildlogic;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import org.gradle.api.DefaultTask;
import org.gradle.api.GradleException;
import org.gradle.api.file.ConfigurableFileCollection;
import org.gradle.api.file.RegularFileProperty;
import org.gradle.api.provider.Property;
import org.gradle.api.tasks.Input;
import org.gradle.api.tasks.InputFile;
import org.gradle.api.tasks.InputFiles;
import org.gradle.api.tasks.OutputFile;
import org.gradle.api.tasks.PathSensitive;
import org.gradle.api.tasks.PathSensitivity;
import org.gradle.api.tasks.TaskAction;

public abstract class VerifyRepositoryFileShapeTask extends DefaultTask {
  @InputFiles
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract ConfigurableFileCollection getSourceFiles();

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
  void verifyRepositoryFileShape() throws IOException {
    Path repositoryRoot = Path.of(getRepositoryRootPath().get()).toAbsolutePath().normalize();
    Path policyPath = getPolicyFile().get().getAsFile().toPath().toAbsolutePath().normalize();
    Path reportPath = getReportFile().get().getAsFile().toPath().toAbsolutePath().normalize();
    JavaSourceShapePolicy policy = JavaSourceShapePolicy.load(policyPath);
    JavaSourceShapePolicy.Rule defaultRule = policy.rules().getLast();
    RepositoryFileShapeAnalyzer analyzer = new RepositoryFileShapeAnalyzer();
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
        throw new GradleException(
            "Repository control-plane shape policy must use PREFIX or DEFAULT rules only: "
                + rule.path());
      }
      if (rule.kind() == JavaSourceShapePolicy.MatchKind.PREFIX
          && relativePaths.stream().noneMatch(rule::matches)) {
        policyIssues.add("PREFIX rule matches no repo-owned control-plane files: " + rule.path());
      }
    }

    for (int index = 0; index < sourceFiles.size(); index++) {
      Path sourceFile = sourceFiles.get(index);
      String relativePath = relativePaths.get(index);
      List<JavaSourceShapePolicy.Rule> matchingRules = policy.matchingRules(relativePath);
      if (matchingRules.isEmpty()) {
        policyIssues.add("No control-plane shape rule matches " + relativePath);
        continue;
      }

      JavaSourceShapePolicy.Rule matchedRule = matchingRules.getFirst();
      SourceShapeMetrics metrics = analyzer.analyze(sourceFile);
      if (matchedRule.kind() == JavaSourceShapePolicy.MatchKind.PREFIX) {
        familyMetrics.computeIfAbsent(matchedRule, unused -> new FamilyMetrics()).include(metrics);
      }
      List<String> exceededMetrics = SourceShapeBudgetSupport.exceededMetrics(metrics, matchedRule);
      if (!exceededMetrics.isEmpty()) {
        violations.add(new Violation(relativePath, matchedRule, exceededMetrics));
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
    List<Path> files = new ArrayList<>();
    for (File entry : getSourceFiles().getFiles()) {
      if (entry.isFile()) {
        files.add(entry.toPath());
        continue;
      }
      if (!entry.isDirectory()) {
        continue;
      }
      try (var walked = Files.walk(entry.toPath())) {
        walked.filter(Files::isRegularFile).sorted().forEach(files::add);
      }
    }
    files.sort(Comparator.naturalOrder());
    return files;
  }

  private String buildFailureMessage(
      Path reportPath, List<String> policyIssues, List<Violation> violations) {
    StringBuilder message = new StringBuilder();
    message
        .append("Repository control-plane shape verification failed.")
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
    }
    return message.toString();
  }

  private String normalizePath(Path relativePath) {
    return relativePath.toString().replace(File.separatorChar, '/');
  }

  private record Violation(
      String relativePath, JavaSourceShapePolicy.Rule rule, List<String> exceededMetrics) {}

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
