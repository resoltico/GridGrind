package dev.erst.gridgrind.buildlogic;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.gradle.api.GradleException;

final class RepositoryFileShapeAnalyzer {
  private static final Pattern KOTLIN_IMPORT_PATTERN = Pattern.compile("^\\s*import\\s+.+$");
  private static final Pattern KOTLIN_TYPE_PATTERN =
      Pattern.compile(
          "^\\s*(?:public|private|internal|protected)?\\s*(?:data\\s+|sealed\\s+|enum\\s+|value\\s+)?(?:class|interface|object)\\s+\\w+.*$");
  private static final Pattern KOTLIN_FUNCTION_PATTERN =
      Pattern.compile(
          "^\\s*(?:public|private|internal|protected)?\\s*(?:suspend\\s+)?fun\\s+(?:<[^>]+>\\s*)?\\w+.*$");
  private static final Pattern KOTLIN_PRIVATE_FUNCTION_PATTERN =
      Pattern.compile(
          "^\\s*(?:private|internal|protected)\\s+(?:suspend\\s+)?fun\\s+(?:<[^>]+>\\s*)?\\w+.*$");
  private static final Pattern KOTLIN_PROPERTY_PATTERN =
      Pattern.compile("^\\s*(?:public|private|internal|protected)?\\s*(?:lateinit\\s+)?(?:val|var)\\s+\\w+.*$");
  private static final Pattern SHELL_INCLUDE_PATTERN =
      Pattern.compile("^\\s*(?:source|\\.)\\s+.+$");
  private static final Pattern SHELL_FUNCTION_PATTERN =
      Pattern.compile("^\\s*(?:function\\s+)?[A-Za-z_][A-Za-z0-9_]*\\s*\\(\\)\\s*\\{\\s*$");
  private static final Pattern SHELL_ASSIGNMENT_PATTERN =
      Pattern.compile(
          "^\\s*(?:local\\s+|readonly\\s+|export\\s+)?[A-Za-z_][A-Za-z0-9_]*=.*$");
  private static final Pattern SHELL_CASE_ARM_PATTERN = Pattern.compile(";;");
  private static final Pattern MARKDOWN_HEADING_PATTERN = Pattern.compile("^#{1,6}\\s+.+$");

  SourceShapeMetrics analyze(Path sourceFile) throws IOException {
    String fileName = sourceFile.getFileName().toString();
    List<String> lines = Files.readAllLines(sourceFile);
    if (fileName.endsWith(".kt") || fileName.endsWith(".kts")) {
      return analyzeKotlin(lines);
    }
    if (fileName.endsWith(".sh")) {
      return analyzeShell(lines);
    }
    if (fileName.endsWith(".md")) {
      return analyzeMarkdown(lines);
    }
    throw new GradleException(
        "Unsupported repository control-plane shape surface: "
            + sourceFile.toAbsolutePath().normalize());
  }

  private SourceShapeMetrics analyzeKotlin(List<String> lines) {
    int importCount = 0;
    int typeCount = 0;
    int methodCount = 0;
    int publicMethodCount = 0;
    int fieldCount = 0;
    for (String line : lines) {
      if (KOTLIN_IMPORT_PATTERN.matcher(line).matches()) {
        importCount++;
      }
      if (KOTLIN_TYPE_PATTERN.matcher(line).matches()) {
        typeCount++;
      }
      if (KOTLIN_FUNCTION_PATTERN.matcher(line).matches()) {
        methodCount++;
        if (!KOTLIN_PRIVATE_FUNCTION_PATTERN.matcher(line).matches()) {
          publicMethodCount++;
        }
      }
      if (KOTLIN_PROPERTY_PATTERN.matcher(line).matches()) {
        fieldCount++;
      }
    }
    return new SourceShapeMetrics(
        lines.size(), importCount, typeCount, 0, methodCount, publicMethodCount, fieldCount, 0, 0);
  }

  private SourceShapeMetrics analyzeShell(List<String> lines) {
    int importCount = 0;
    int methodCount = 0;
    int fieldCount = 0;
    int switchCount = 0;
    int maxSwitchArms = 0;
    int currentCaseArms = 0;
    boolean insideCase = false;

    for (String line : lines) {
      if (SHELL_INCLUDE_PATTERN.matcher(line).matches()) {
        importCount++;
      }
      if (SHELL_FUNCTION_PATTERN.matcher(line).matches()) {
        methodCount++;
      }
      if (SHELL_ASSIGNMENT_PATTERN.matcher(line).matches()) {
        fieldCount++;
      }

      String trimmed = line.trim();
      if (trimmed.startsWith("case ")) {
        switchCount++;
        insideCase = true;
        currentCaseArms = 0;
      }
      if (insideCase) {
        Matcher caseArmMatcher = SHELL_CASE_ARM_PATTERN.matcher(trimmed);
        while (caseArmMatcher.find()) {
          currentCaseArms++;
        }
        if (trimmed.equals("esac")) {
          maxSwitchArms = Math.max(maxSwitchArms, currentCaseArms);
          insideCase = false;
        }
      }
    }

    return new SourceShapeMetrics(
        lines.size(), importCount, 0, 0, methodCount, methodCount, fieldCount, switchCount, maxSwitchArms);
  }

  private SourceShapeMetrics analyzeMarkdown(List<String> lines) {
    int headingCount = 0;
    for (String line : lines) {
      if (MARKDOWN_HEADING_PATTERN.matcher(line).matches()) {
        headingCount++;
      }
    }
    return new SourceShapeMetrics(lines.size(), 0, headingCount, 0, 0, 0, 0, 0, 0);
  }
}
