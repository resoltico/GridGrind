package dev.erst.gridgrind.buildlogic;

import java.io.BufferedWriter;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
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
    JavaSourceShapeAnalyzer analyzer = new JavaSourceShapeAnalyzer();
    LocalDate today = LocalDate.now();

    List<Path> sourceFiles = collectSourceFiles();
    List<String> relativePaths =
        sourceFiles.stream().map(repositoryRoot::relativize).map(this::normalizePath).toList();
    List<String> policyIssues = new ArrayList<>();
    List<Violation> violations = new ArrayList<>();
    List<ReportRow> reportRows = new ArrayList<>();

    for (JavaSourceShapePolicy.Rule rule : policy.rules()) {
      if (rule.kind() == JavaSourceShapePolicy.MatchKind.EXACT) {
        if (!Files.isRegularFile(repositoryRoot.resolve(rule.path()))) {
          policyIssues.add("EXACT rule points at a missing file: " + rule.path());
        }
      }
      if (rule.kind() == JavaSourceShapePolicy.MatchKind.PREFIX
          && relativePaths.stream().noneMatch(rule::matches)) {
        policyIssues.add("PREFIX rule matches no production source files: " + rule.path());
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
      JavaSourceShapeAnalyzer.Metrics metrics = analyzer.analyze(sourceFile, javaRelease);
      List<String> exceededMetrics = exceededMetrics(metrics, matchedRule);
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

    writeReport(reportPath, reportRows);

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

  private void writeReport(Path reportPath, List<ReportRow> reportRows) throws IOException {
    Files.createDirectories(reportPath.getParent());
    List<ReportRow> sortedRows =
        reportRows.stream()
            .sorted(Comparator.comparingDouble(ReportRow::riskScore).reversed())
            .toList();
    try (BufferedWriter writer = Files.newBufferedWriter(reportPath, StandardCharsets.UTF_8)) {
      writer.write(
          "path\tkind\trole\tlines\tmaxLines\tmethods\tmaxMethods\tpublicMethods"
              + "\tmaxPublicMethods\timports\tmaxImports\tfields\tmaxFields\tswitches"
              + "\tmaxSwitches\tmaxSwitchArms\tlimitMaxSwitchArms\ttopLevelTypes"
              + "\tnestedTypes\trisk\towner\treviewExpiresOn\tsplitTrigger");
      writer.newLine();
      for (ReportRow row : sortedRows) {
        writer.write(row.toTsv());
        writer.newLine();
      }
    }
  }

  private List<String> exceededMetrics(
      JavaSourceShapeAnalyzer.Metrics metrics, JavaSourceShapePolicy.Rule rule) {
    List<String> exceededMetrics = new ArrayList<>();
    if (exceeds(metrics.lineCount(), rule.maxLines())) {
      exceededMetrics.add("lines=" + metrics.lineCount() + ">" + rule.maxLines());
    }
    if (exceeds(metrics.methodCount(), rule.maxMethods())) {
      exceededMetrics.add("methods=" + metrics.methodCount() + ">" + rule.maxMethods());
    }
    if (exceeds(metrics.publicMethodCount(), rule.maxPublicMethods())) {
      exceededMetrics.add(
          "publicMethods=" + metrics.publicMethodCount() + ">" + rule.maxPublicMethods());
    }
    if (exceeds(metrics.importCount(), rule.maxImports())) {
      exceededMetrics.add("imports=" + metrics.importCount() + ">" + rule.maxImports());
    }
    if (exceeds(metrics.fieldCount(), rule.maxFields())) {
      exceededMetrics.add("fields=" + metrics.fieldCount() + ">" + rule.maxFields());
    }
    if (exceeds(metrics.switchCount(), rule.maxSwitches())) {
      exceededMetrics.add("switches=" + metrics.switchCount() + ">" + rule.maxSwitches());
    }
    if (exceeds(metrics.maxSwitchArms(), rule.maxSwitchArms())) {
      exceededMetrics.add(
          "maxSwitchArms=" + metrics.maxSwitchArms() + ">" + rule.maxSwitchArms());
    }
    return List.copyOf(exceededMetrics);
  }

  private boolean exceeds(long value, Integer limit) {
    return limit != null && value > limit;
  }

  private String normalizePath(Path relativePath) {
    return relativePath.toString().replace(File.separatorChar, '/');
  }

  private record Violation(
      String relativePath,
      JavaSourceShapePolicy.Rule rule,
      JavaSourceShapeAnalyzer.Metrics metrics,
      List<String> exceededMetrics) {}

  private record ReportRow(
      String relativePath,
      JavaSourceShapePolicy.Rule rule,
      JavaSourceShapeAnalyzer.Metrics metrics) {
    private double riskScore() {
      return ratio(metrics.lineCount(), rule.maxLines())
          + ratio(metrics.methodCount(), rule.maxMethods())
          + ratio(metrics.publicMethodCount(), rule.maxPublicMethods())
          + ratio(metrics.importCount(), rule.maxImports())
          + ratio(metrics.fieldCount(), rule.maxFields())
          + ratio(metrics.switchCount(), rule.maxSwitches())
          + ratio(metrics.maxSwitchArms(), rule.maxSwitchArms());
    }

    private String toTsv() {
      return relativePath
          + '\t'
          + rule.kind()
          + '\t'
          + rule.role()
          + '\t'
          + metrics.lineCount()
          + '\t'
          + limit(rule.maxLines())
          + '\t'
          + metrics.methodCount()
          + '\t'
          + limit(rule.maxMethods())
          + '\t'
          + metrics.publicMethodCount()
          + '\t'
          + limit(rule.maxPublicMethods())
          + '\t'
          + metrics.importCount()
          + '\t'
          + limit(rule.maxImports())
          + '\t'
          + metrics.fieldCount()
          + '\t'
          + limit(rule.maxFields())
          + '\t'
          + metrics.switchCount()
          + '\t'
          + limit(rule.maxSwitches())
          + '\t'
          + metrics.maxSwitchArms()
          + '\t'
          + limit(rule.maxSwitchArms())
          + '\t'
          + metrics.topLevelTypeCount()
          + '\t'
          + metrics.nestedTypeCount()
          + '\t'
          + String.format(Locale.ROOT, "%.2f", riskScore())
          + '\t'
          + rule.owner()
          + '\t'
          + limit(rule.reviewExpiresOn())
          + '\t'
          + limit(rule.splitTrigger());
    }

    private double ratio(long value, Integer limit) {
      return limit == null || limit == 0 ? 0.0d : (double) value / (double) limit;
    }

    private String limit(Integer limit) {
      return limit == null ? "-" : Integer.toString(limit);
    }

    private String limit(LocalDate limit) {
      return limit == null ? "-" : limit.toString();
    }

    private String limit(String limit) {
      return limit == null ? "-" : limit;
    }
  }
}
