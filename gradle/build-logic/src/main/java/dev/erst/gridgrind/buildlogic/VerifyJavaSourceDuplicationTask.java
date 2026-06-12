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
import net.sourceforge.pmd.cpd.CPDConfiguration;
import net.sourceforge.pmd.cpd.CPDReport;
import net.sourceforge.pmd.cpd.CpdAnalysis;
import net.sourceforge.pmd.cpd.Mark;
import net.sourceforge.pmd.cpd.Match;
import net.sourceforge.pmd.lang.LanguageRegistry;
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

public abstract class VerifyJavaSourceDuplicationTask extends DefaultTask {
  private static final int MINIMUM_TOKEN_SPAN = 320;
  private static final int MAX_SUMMARY_MATCHES = 10;

  @InputFiles
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract ConfigurableFileCollection getSourceRoots();

  @Input
  public abstract Property<String> getRepositoryRootPath();

  @InputFile
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract RegularFileProperty getPolicyFile();

  @OutputFile
  public abstract RegularFileProperty getReportFile();

  @TaskAction
  void verifyDuplication() throws IOException {
    Path repositoryRoot = Path.of(getRepositoryRootPath().get()).toAbsolutePath().normalize();
    Path policyPath = getPolicyFile().get().getAsFile().toPath().toAbsolutePath().normalize();
    Path reportPath = getReportFile().get().getAsFile().toPath().toAbsolutePath().normalize();
    JavaSourceShapePolicy policy = JavaSourceShapePolicy.load(policyPath);
    List<Path> sourceFiles = collectSourceFiles(repositoryRoot, policy);
    if (sourceFiles.isEmpty()) {
      Files.deleteIfExists(reportPath);
      return;
    }

    Files.createDirectories(reportPath.getParent());

    CPDConfiguration configuration = new CPDConfiguration(LanguageRegistry.CPD);
    configuration.setSourceEncoding(StandardCharsets.UTF_8);
    configuration.setMinimumTileSize(MINIMUM_TOKEN_SPAN);
    configuration.setIgnoreIdentifiers(true);
    configuration.setIgnoreLiterals(true);
    configuration.setIgnoreAnnotations(true);
    configuration.setSkipDuplicates(true);

    CPDReport report;
    try (CpdAnalysis analysis = CpdAnalysis.create(configuration)) {
      for (Path sourceFile : sourceFiles) {
        analysis.files().addFile(sourceFile);
      }
      final CPDReport[] reportHolder = new CPDReport[1];
      analysis.performAnalysis(
          duplicationReport -> {
            reportHolder[0] = duplicationReport;
          });
      report = reportHolder[0];
    }

    if (report == null) {
      throw new GradleException(
          "Java source-duplication verification did not produce a CPD report at " + reportPath + ".");
    }
    writeReport(repositoryRoot, reportPath, report);
    if (!report.getProcessingErrors().isEmpty()) {
      throw new GradleException(
          "Java source-duplication verification encountered CPD processing errors. See "
              + reportPath
              + ".");
    }

    List<Match> matches = report.getMatches();
    if (!matches.isEmpty()) {
      throw new GradleException(buildFailureMessage(repositoryRoot, reportPath, matches, report));
    }
  }

  private List<Path> collectSourceFiles(Path repositoryRoot, JavaSourceShapePolicy policy)
      throws IOException {
    List<Path> sourceFiles = new ArrayList<>();
    Set<File> roots = getSourceRoots().getFiles();
    for (File sourceRoot : roots) {
      if (!sourceRoot.isDirectory()) {
        continue;
      }
      try (var walked = Files.walk(sourceRoot.toPath())) {
        walked
            .filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".java"))
            .filter(path -> !path.getFileName().toString().equals("module-info.java"))
            .filter(path -> !path.getFileName().toString().equals("package-info.java"))
            .filter(
                path ->
                    duplicationGuardFor(repositoryRoot, policy, path)
                        == JavaSourceShapePolicy.DuplicationGuard.CHECK)
            .sorted()
            .forEach(sourceFiles::add);
      }
    }
    sourceFiles.sort(Comparator.naturalOrder());
    return sourceFiles;
  }

  private JavaSourceShapePolicy.DuplicationGuard duplicationGuardFor(
      Path repositoryRoot, JavaSourceShapePolicy policy, Path sourceFile) {
    String relativePath =
        repositoryRoot.relativize(sourceFile.toAbsolutePath().normalize()).toString().replace(File.separatorChar, '/');
    List<JavaSourceShapePolicy.Rule> matchingRules = policy.matchingRules(relativePath);
    if (matchingRules.isEmpty()) {
      throw new GradleException("No source-shape rule matches " + relativePath + ".");
    }
    return matchingRules.getFirst().duplicationGuard();
  }

  private void writeReport(Path repositoryRoot, Path reportPath, CPDReport report) throws IOException {
    try (BufferedWriter writer = Files.newBufferedWriter(reportPath, StandardCharsets.UTF_8)) {
      writer.write("tokens\tlines\toccurrences\tlocations");
      writer.newLine();
      for (Match match : report.getMatches().stream().sorted(Match.MATCHES_COMPARATOR.reversed()).toList()) {
        List<String> locations = new ArrayList<>();
        for (Mark mark : match) {
          locations.add(locationText(repositoryRoot, report, mark));
        }
        writer.write(
            match.getTokenCount()
                + "\t"
                + match.getLineCount()
                + "\t"
                + match.getMarkCount()
                + "\t"
                + String.join(" | ", locations));
        writer.newLine();
      }
    }
  }

  private String buildFailureMessage(
      Path repositoryRoot, Path reportPath, List<Match> matches, CPDReport report) {
    StringBuilder message = new StringBuilder();
    message
        .append("Java source-duplication verification failed.")
        .append(System.lineSeparator())
        .append("Report: ")
        .append(reportPath)
        .append(System.lineSeparator())
        .append("Matches:")
        .append(System.lineSeparator());
    for (Match match : matches.stream().sorted(Match.MATCHES_COMPARATOR.reversed()).limit(MAX_SUMMARY_MATCHES).toList()) {
      message
          .append(" - ")
          .append(match.getTokenCount())
          .append(" tokens / ")
          .append(match.getLineCount())
          .append(" lines across ");
      List<String> locations = new ArrayList<>();
      for (Mark mark : match) {
        locations.add(locationText(repositoryRoot, report, mark));
      }
      message.append(String.join(", ", locations)).append(System.lineSeparator());
    }
    if (matches.size() > MAX_SUMMARY_MATCHES) {
          message
          .append(" - ... and ")
          .append(matches.size() - MAX_SUMMARY_MATCHES)
          .append(" more duplicate blocks. See the TSV report for the full list.")
          .append(System.lineSeparator());
    }
    return message.toString();
  }

  private String locationText(Path repositoryRoot, CPDReport report, Mark mark) {
    String displayName = report.getDisplayName(mark.getLocation().getFileId());
    Path displayPath = Path.of(displayName).toAbsolutePath().normalize();
    String relativePath;
    if (displayPath.startsWith(repositoryRoot)) {
      relativePath = repositoryRoot.relativize(displayPath).toString();
    } else {
      relativePath = displayPath.toString();
    }
    return relativePath.replace(File.separatorChar, '/') + ":" + mark.getLocation().getStartLine();
  }
}
