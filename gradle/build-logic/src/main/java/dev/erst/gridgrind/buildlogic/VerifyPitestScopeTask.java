package dev.erst.gridgrind.buildlogic;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Comparator;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;
import org.gradle.api.DefaultTask;
import org.gradle.api.GradleException;
import org.gradle.api.file.ConfigurableFileCollection;
import org.gradle.api.file.RegularFileProperty;
import org.gradle.api.provider.SetProperty;
import org.gradle.api.tasks.Input;
import org.gradle.api.tasks.InputFiles;
import org.gradle.api.tasks.OutputFile;
import org.gradle.api.tasks.PathSensitive;
import org.gradle.api.tasks.PathSensitivity;
import org.gradle.api.tasks.TaskAction;

/** Verifies that every configured PIT class and test pattern resolves to compiled bytecode. */
public abstract class VerifyPitestScopeTask extends DefaultTask {
  @Input
  public abstract SetProperty<String> getTargetClassPatterns();

  @Input
  public abstract SetProperty<String> getTargetTestPatterns();

  @InputFiles
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract ConfigurableFileCollection getProductionClassDirectories();

  @InputFiles
  @PathSensitive(PathSensitivity.RELATIVE)
  public abstract ConfigurableFileCollection getTestClassDirectories();

  @OutputFile
  public abstract RegularFileProperty getReportFile();

  @TaskAction
  void verifyScope() throws IOException {
    Set<String> productionClasses = classNames(getProductionClassDirectories().getFiles());
    Set<String> testClasses = classNames(getTestClassDirectories().getFiles());
    List<ScopeEntry> entries = new ArrayList<>();
    List<String> violations = new ArrayList<>();
    verifyPatterns("production class", getTargetClassPatterns().get(), productionClasses, entries, violations);
    verifyPatterns("test class", getTargetTestPatterns().get(), testClasses, entries, violations);
    writeReport(entries);
    if (!violations.isEmpty()) {
      throw new GradleException(
          "PIT scope verification failed. Every configured pattern must resolve to compiled bytecode:\n - "
              + String.join("\n - ", violations));
    }
  }

  private static void verifyPatterns(
      String kind,
      Set<String> patterns,
      Collection<String> availableClasses,
      List<ScopeEntry> entries,
      List<String> violations) {
    if (patterns.isEmpty()) {
      violations.add("no " + kind + " patterns are configured");
      return;
    }
    for (String pattern : patterns.stream().sorted().toList()) {
      List<String> matches = MutationScopePatternSupport.matchingClassNames(availableClasses, pattern);
      entries.add(new ScopeEntry(kind, pattern, matches));
      if (matches.isEmpty()) {
        violations.add(kind + " pattern '" + pattern + "' matched no compiled class");
      }
    }
  }

  private static Set<String> classNames(Collection<File> classDirectories) throws IOException {
    Set<String> classNames = new TreeSet<>();
    for (File classDirectory : classDirectories) {
      if (!classDirectory.isDirectory()) {
        continue;
      }
      Path root = classDirectory.toPath();
      try (var paths = Files.walk(root)) {
        paths
            .filter(path -> Files.isRegularFile(path) && path.toString().endsWith(".class"))
            .map(path -> className(root, path))
            .filter(className -> !className.equals("module-info"))
            .forEach(classNames::add);
      }
    }
    return classNames;
  }

  private static String className(Path root, Path classFile) {
    String relativePath = root.relativize(classFile).toString().replace(File.separatorChar, '/');
    return relativePath.substring(0, relativePath.length() - ".class".length()).replace('/', '.');
  }

  private void writeReport(List<ScopeEntry> entries) throws IOException {
    Path reportPath = getReportFile().get().getAsFile().toPath();
    Files.createDirectories(reportPath.getParent());
    List<String> lines = new ArrayList<>();
    lines.add("kind\tpattern\tmatchedClasses");
    entries.stream()
        .sorted(Comparator.comparing(ScopeEntry::kind).thenComparing(ScopeEntry::pattern))
        .forEach(
            entry ->
                lines.add(
                    entry.kind()
                        + "\t"
                        + entry.pattern()
                        + "\t"
                        + String.join(",", entry.matches())));
    Files.write(reportPath, lines);
  }

  private record ScopeEntry(String kind, String pattern, List<String> matches) {}
}
