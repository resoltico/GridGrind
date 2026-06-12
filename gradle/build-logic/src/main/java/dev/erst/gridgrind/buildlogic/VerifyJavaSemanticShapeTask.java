package dev.erst.gridgrind.buildlogic;

import java.io.BufferedWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import javax.xml.parsers.DocumentBuilderFactory;
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
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

public abstract class VerifyJavaSemanticShapeTask extends DefaultTask {
  @InputFiles
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract ConfigurableFileCollection getReportFiles();

  @Input
  public abstract Property<String> getRepositoryRootPath();

  @Input
  public abstract Property<String> getReviewDate();

  @InputFile
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract RegularFileProperty getPolicyFile();

  @OutputFile
  public abstract RegularFileProperty getOutputReportFile();

  @TaskAction
  void verifySemanticShape() throws Exception {
    Path repositoryRoot = Path.of(getRepositoryRootPath().get()).toAbsolutePath().normalize();
    Path policyPath = getPolicyFile().get().getAsFile().toPath().toAbsolutePath().normalize();
    Path outputReportPath =
        getOutputReportFile().get().getAsFile().toPath().toAbsolutePath().normalize();
    LocalDate today = LocalDate.parse(getReviewDate().get());
    JavaSemanticShapePolicy policy = JavaSemanticShapePolicy.load(policyPath);

    List<Violation> violations = collectViolations(repositoryRoot);
    List<String> policyIssues = new ArrayList<>();
    List<ViolationRecord> reportRows = new ArrayList<>();
    java.util.Map<JavaSemanticShapePolicy.Rule, Integer> matchedViolationCounts =
        new java.util.LinkedHashMap<>();

    for (JavaSemanticShapePolicy.Rule rule : policy.rules()) {
      Path sourcePath = repositoryRoot.resolve(rule.path());
      if (rule.kind() == JavaSourceShapePolicy.MatchKind.EXACT && !Files.isRegularFile(sourcePath)) {
        policyIssues.add("EXACT semantic-shape rule points at a missing file: " + rule.path());
      }
      if (rule.kind() == JavaSourceShapePolicy.MatchKind.PREFIX
          && violations.stream().noneMatch(violation -> rule.matches(violation.relativePath()))) {
        policyIssues.add(
            "PREFIX semantic-shape rule matches no current semantic PMD findings: " + rule.path());
      }
      if (today.isAfter(rule.reviewExpiresOn())) {
        policyIssues.add(
            "Semantic-shape review expired for "
                + rule.path()
                + " on "
                + rule.reviewExpiresOn()
                + ".");
      }
    }

    for (Violation violation : violations) {
      JavaSemanticShapePolicy.Rule matchingRule = policy.matchingRule(violation.relativePath());
      if (matchingRule == null) {
        reportRows.add(ViolationRecord.unreviewed(violation));
        continue;
      }
      boolean accepted = matchingRule.allows(violation.rule());
      if (accepted) {
        matchedViolationCounts.merge(matchingRule, 1, Integer::sum);
      }
      reportRows.add(ViolationRecord.reviewed(violation, matchingRule, accepted));
    }

    for (JavaSemanticShapePolicy.Rule rule : policy.rules()) {
      if (!matchedViolationCounts.containsKey(rule)) {
        policyIssues.add(
            "Semantic-shape rule matches no current PMD findings and should be removed: "
                + rule.path()
                + " ["
                + String.join(",", rule.allowedRules())
                + "]");
      }
    }

    Files.createDirectories(outputReportPath.getParent());
    writeReport(outputReportPath, reportRows);

    List<ViolationRecord> unresolved =
        reportRows.stream().filter(row -> !row.accepted()).sorted().toList();
    if (!policyIssues.isEmpty() || !unresolved.isEmpty()) {
      throw new GradleException(buildFailureMessage(outputReportPath, policyIssues, unresolved));
    }
  }

  private List<Violation> collectViolations(Path repositoryRoot) throws Exception {
    List<Violation> violations = new ArrayList<>();
    for (var reportFile : getReportFiles().getFiles()) {
      if (!reportFile.isFile()) {
        continue;
      }
      violations.addAll(parseViolations(repositoryRoot, reportFile.toPath()));
    }
    violations.sort(Comparator.naturalOrder());
    return List.copyOf(violations);
  }

  private List<Violation> parseViolations(Path repositoryRoot, Path reportPath) throws Exception {
    List<Violation> violations = new ArrayList<>();
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    factory.setNamespaceAware(true);
    Document document = factory.newDocumentBuilder().parse(reportPath.toFile());
    Element root = document.getDocumentElement();
    forEachChild(
        root,
        "file",
        fileElement -> {
          String relativePath = normalizeReportPath(repositoryRoot, fileElement.getAttribute("name"));
          forEachChild(
              fileElement,
              "violation",
              violationElement ->
                  violations.add(
                      new Violation(
                          relativePath,
                          violationElement.getAttribute("rule"),
                          violationElement.getAttribute("class"),
                          emptyToNull(violationElement.getAttribute("method")),
                          parseInt(violationElement.getAttribute("beginline")),
                          parseInt(violationElement.getAttribute("priority")),
                          normalizeMessage(violationElement.getTextContent()))));
        });
    return violations;
  }

  private void forEachChild(
      Element parent, String localName, java.util.function.Consumer<Element> consumer) {
    for (Node child = parent.getFirstChild(); child != null; child = child.getNextSibling()) {
      if (child instanceof Element element && localName.equals(element.getLocalName())) {
        consumer.accept(element);
      }
    }
  }

  private String normalizeReportPath(Path repositoryRoot, String absolutePath) {
    Path normalized = Path.of(absolutePath).toAbsolutePath().normalize();
    if (!normalized.startsWith(repositoryRoot)) {
      throw new GradleException(
          "Semantic-shape PMD report referenced a file outside the repository root: "
              + absolutePath);
    }
    return repositoryRoot
        .relativize(normalized)
        .toString()
        .replace(FileSystems.getDefault().getSeparator(), "/");
  }

  private String buildFailureMessage(
      Path outputReportPath, List<String> policyIssues, List<ViolationRecord> unresolved) {
    StringBuilder message = new StringBuilder();
    message
        .append("Java semantic-shape verification failed.")
        .append(System.lineSeparator())
        .append("Report: ")
        .append(outputReportPath)
        .append(System.lineSeparator());
    if (!policyIssues.isEmpty()) {
      message.append("Policy issues:").append(System.lineSeparator());
      for (String issue : policyIssues) {
        message.append(" - ").append(issue).append(System.lineSeparator());
      }
    }
    if (!unresolved.isEmpty()) {
      message.append("Unreviewed semantic-shape findings:").append(System.lineSeparator());
      for (ViolationRecord violation : unresolved.stream().limit(30).toList()) {
        message
            .append(" - ")
            .append(violation.relativePath())
            .append(" [")
            .append(violation.rule())
            .append("]");
        if (violation.methodName() != null) {
          message.append(" ").append(violation.className()).append("#").append(violation.methodName());
        } else {
          message.append(" ").append(violation.className());
        }
        message.append(": ").append(violation.message()).append(System.lineSeparator());
      }
      if (unresolved.size() > 30) {
        message
            .append(" - ... and ")
            .append(unresolved.size() - 30)
            .append(" more. See the TSV report for the full list.")
            .append(System.lineSeparator());
      }
    }
    return message.toString();
  }

  private void writeReport(Path outputReportPath, List<ViolationRecord> reportRows) throws IOException {
    try (BufferedWriter writer =
        Files.newBufferedWriter(outputReportPath, StandardCharsets.UTF_8)) {
      writer.write(
          "path\trule\tclass\tmethod\tbeginLine\tpriority\tstatus\trole\towner\treviewExpiresOn\tsplitTrigger\tmessage");
      writer.newLine();
      for (ViolationRecord reportRow : reportRows.stream().sorted().toList()) {
        writer.write(reportRow.toTsv());
        writer.newLine();
      }
    }
  }

  private static String normalizeMessage(String message) {
    return message == null ? "" : message.trim().replaceAll("\\s+", " ");
  }

  private static Integer parseInt(String rawValue) {
    if (rawValue == null || rawValue.isBlank()) {
      return null;
    }
    return Integer.parseInt(rawValue);
  }

  private static String emptyToNull(String rawValue) {
    return rawValue == null || rawValue.isBlank() ? null : rawValue;
  }

  private record Violation(
      String relativePath,
      String rule,
      String className,
      String methodName,
      Integer beginLine,
      Integer priority,
      String message)
      implements Comparable<Violation> {
    @Override
    public int compareTo(Violation other) {
      return Comparator.comparing(Violation::relativePath)
          .thenComparing(Violation::rule)
          .thenComparing(Violation::className)
          .thenComparing(violation -> violation.methodName() == null ? "" : violation.methodName())
          .thenComparing(
              violation -> violation.beginLine() == null ? Integer.MAX_VALUE : violation.beginLine())
          .compare(this, other);
    }
  }

  private record ViolationRecord(
      String relativePath,
      String rule,
      String className,
      String methodName,
      Integer beginLine,
      Integer priority,
      boolean accepted,
      String role,
      String owner,
      LocalDate reviewExpiresOn,
      String splitTrigger,
      String message)
      implements Comparable<ViolationRecord> {
    private static ViolationRecord unreviewed(Violation violation) {
      return new ViolationRecord(
          violation.relativePath(),
          violation.rule(),
          violation.className(),
          violation.methodName(),
          violation.beginLine(),
          violation.priority(),
          false,
          "-",
          "-",
          null,
          "-",
          violation.message());
    }

    private static ViolationRecord reviewed(
        Violation violation, JavaSemanticShapePolicy.Rule rule, boolean accepted) {
      return new ViolationRecord(
          violation.relativePath(),
          violation.rule(),
          violation.className(),
          violation.methodName(),
          violation.beginLine(),
          violation.priority(),
          accepted,
          rule.role(),
          rule.owner(),
          rule.reviewExpiresOn(),
          rule.splitTrigger(),
          violation.message());
    }

    private String toTsv() {
      return relativePath
          + '\t'
          + rule
          + '\t'
          + className
          + '\t'
          + nullSafe(methodName)
          + '\t'
          + nullSafe(beginLine)
          + '\t'
          + nullSafe(priority)
          + '\t'
          + (accepted ? "REVIEWED" : "UNREVIEWED")
          + '\t'
          + role
          + '\t'
          + owner
          + '\t'
          + nullSafe(reviewExpiresOn)
          + '\t'
          + splitTrigger
          + '\t'
          + message;
    }

    private static String nullSafe(Object value) {
      return value == null ? "-" : value.toString();
    }

    @Override
    public int compareTo(ViolationRecord other) {
      return Comparator.comparing(ViolationRecord::relativePath)
          .thenComparing(ViolationRecord::rule)
          .thenComparing(ViolationRecord::className)
          .thenComparing(record -> record.methodName() == null ? "" : record.methodName())
          .thenComparing(record -> record.beginLine() == null ? Integer.MAX_VALUE : record.beginLine())
          .thenComparing(record -> record.accepted() ? 1 : 0)
          .compare(this, other);
    }
  }
}
