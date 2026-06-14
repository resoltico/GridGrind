package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

class RepositoryFileShapeAnalyzerTest {
  @Test
  void analyzesKotlinShellAndMarkdownControlPlaneSurfaces() throws IOException {
    RepositoryFileShapeAnalyzer analyzer = new RepositoryFileShapeAnalyzer();

    Path kotlinFile = Files.createTempFile("gridgrind-control-plane", ".kt");
    Files.writeString(
        kotlinFile,
        """
        import org.gradle.api.Project

        class SamplePlugin {
          private val state = 1

          fun apply(project: Project) {}

          private fun helper() {}
        }
        """);

    Path shellFile = Files.createTempFile("gridgrind-control-plane", ".sh");
    Files.writeString(
        shellFile,
        """
        #!/usr/bin/env bash
        source helper.sh
        readonly NAME=value

        run_gate() {
          case "$1" in
            alpha) ;;
            beta) ;;
          esac
        }
        """);

    Path markdownFile = Files.createTempFile("gridgrind-control-plane", ".md");
    Files.writeString(
        markdownFile,
        """
        # Title

        ## Detail
        """);

    SourceShapeMetrics kotlinMetrics = analyzer.analyze(kotlinFile);
    SourceShapeMetrics shellMetrics = analyzer.analyze(shellFile);
    SourceShapeMetrics markdownMetrics = analyzer.analyze(markdownFile);

    assertEquals(1, kotlinMetrics.importCount());
    assertEquals(1, kotlinMetrics.topLevelTypeCount());
    assertEquals(2, kotlinMetrics.methodCount());
    assertEquals(1, kotlinMetrics.publicMethodCount());
    assertEquals(1, kotlinMetrics.fieldCount());

    assertEquals(1, shellMetrics.importCount());
    assertEquals(1, shellMetrics.methodCount());
    assertEquals(1, shellMetrics.fieldCount());
    assertEquals(1, shellMetrics.switchCount());
    assertEquals(2, shellMetrics.maxSwitchArms());

    assertEquals(2, markdownMetrics.topLevelTypeCount());
    assertEquals(0, markdownMetrics.methodCount());
  }
}
