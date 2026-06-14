package dev.erst.gridgrind.buildlogic;

import java.io.BufferedWriter;
import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Set;
import org.gradle.api.DefaultTask;
import org.gradle.api.GradleException;
import org.gradle.api.file.ConfigurableFileCollection;
import org.gradle.api.file.RegularFileProperty;
import org.gradle.api.provider.Property;
import org.gradle.api.tasks.Input;
import org.gradle.api.tasks.InputFiles;
import org.gradle.api.tasks.OutputFile;
import org.gradle.api.tasks.PathSensitive;
import org.gradle.api.tasks.PathSensitivity;
import org.gradle.api.tasks.TaskAction;

public abstract class VerifyForbiddenJavaUnionShapesTask extends DefaultTask {
  @InputFiles
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract ConfigurableFileCollection getSourceRoots();

  @Input
  public abstract Property<Integer> getJavaRelease();

  @Input
  public abstract Property<String> getRepositoryRootPath();

  @OutputFile
  public abstract RegularFileProperty getReportFile();

  @TaskAction
  void verifyForbiddenShapes() throws IOException {
    Path repositoryRoot = Path.of(getRepositoryRootPath().get()).toAbsolutePath().normalize();
    Path reportPath = getReportFile().get().getAsFile().toPath().toAbsolutePath().normalize();
    JavaForbiddenTopLevelShapeAnalyzer analyzer = new JavaForbiddenTopLevelShapeAnalyzer();
    List<ReportRow> violations = new ArrayList<>();

    for (Path sourceFile : collectProductionJavaFiles()) {
      String relativePath = normalizePath(repositoryRoot.relativize(sourceFile.toAbsolutePath().normalize()));
      for (JavaForbiddenTopLevelShapeAnalyzer.Violation violation :
          analyzer.analyze(sourceFile, getJavaRelease().get())) {
        violations.add(new ReportRow(relativePath, violation));
      }
    }

    writeReport(reportPath, violations);
    if (!violations.isEmpty()) {
      throw new GradleException(buildFailureMessage(reportPath, violations));
    }
  }

  private List<Path> collectProductionJavaFiles() throws IOException {
    List<Path> sourceFiles = new ArrayList<>();
    Set<File> roots = getSourceRoots().getFiles();
    for (File sourceRoot : roots) {
      if (!sourceRoot.isDirectory()) {
        continue;
      }
      try (var walked = Files.walk(sourceRoot.toPath())) {
        walked
            .filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".java"))
            .filter(this::isProductionJavaPath)
            .sorted()
            .forEach(sourceFiles::add);
      }
    }
    sourceFiles.sort(Comparator.naturalOrder());
    return sourceFiles;
  }

  private boolean isProductionJavaPath(Path path) {
    String normalized = normalizePath(path);
    return normalized.contains("/src/main/java/");
  }

  private void writeReport(Path reportPath, List<ReportRow> rows) throws IOException {
    Files.createDirectories(reportPath.getParent());
    try (BufferedWriter writer = Files.newBufferedWriter(reportPath, StandardCharsets.UTF_8)) {
      writer.write("path\ttypeKind\ttypeName\tstateSlots\toptionalStateSlots\tdiscriminatorSlots\treason");
      writer.newLine();
      for (ReportRow row : rows) {
        writer.write(row.toTsv());
        writer.newLine();
      }
    }
  }

  private String buildFailureMessage(Path reportPath, List<ReportRow> rows) {
    StringBuilder message = new StringBuilder();
    message
        .append("Forbidden tagged-union or god-record shapes were found.")
        .append(System.lineSeparator())
        .append("Report: ")
        .append(reportPath)
        .append(System.lineSeparator())
        .append("Violations:")
        .append(System.lineSeparator());
    for (ReportRow row : rows.stream().limit(20).toList()) {
      message
          .append(" - ")
          .append(row.relativePath())
          .append(" :: ")
          .append(row.violation().typeKind())
          .append(" ")
          .append(row.violation().typeName())
          .append(" :: ")
          .append(row.violation().reason())
          .append(System.lineSeparator());
    }
    if (rows.size() > 20) {
      message
          .append(" - ... and ")
          .append(rows.size() - 20)
          .append(" more. See the TSV report for the full list.")
          .append(System.lineSeparator());
    }
    return message.toString();
  }

  private String normalizePath(Path path) {
    return path.toString().replace(File.separatorChar, '/');
  }

  private record ReportRow(
      String relativePath, JavaForbiddenTopLevelShapeAnalyzer.Violation violation) {
    private String toTsv() {
      return relativePath
          + '\t'
          + violation.typeKind()
          + '\t'
          + violation.typeName()
          + '\t'
          + violation.stateSlots()
          + '\t'
          + violation.optionalStateSlots()
          + '\t'
          + String.join(",", violation.discriminatorSlots())
          + '\t'
          + violation.reason();
    }
  }
}
