package dev.erst.gridgrind.buildlogic;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import org.gradle.testfixtures.ProjectBuilder;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Checks report retention without executing product tests or changing dependency caches. */
class GridGrindReportRetentionTest {
  @TempDir Path root;

  @Test
  void keepsThreeNewestReportCapturesAndLeavesOtherReportsAlone() throws Exception {
    var project = ProjectBuilder.builder().withProjectDir(root.toFile()).build();
    Path reports = root.resolve("build/reports/configuration-cache");
    for (int index = 0; index < 5; index++) {
      Path report = reports.resolve("key-" + index).resolve("capture/configuration-cache-report.html");
      Files.createDirectories(report.getParent());
      Files.writeString(report, "configuration diagnostics " + index);
      Files.setLastModifiedTime(report, FileTime.fromMillis(1_000L * index));
    }
    Path captureNote = reports.resolve("key-0/capture/keep.txt");
    Files.writeString(captureNote, "unrelated capture note");
    Path other = root.resolve("build/reports/tests/retained.html");
    Files.createDirectories(other.getParent());
    Files.writeString(other, "test failure evidence");

    var task = GridGrindReportRetentionKt.registerGridGrindReportRetention(project).get();
    task.getActions().forEach(action -> action.execute(task));

    assertFalse(Files.exists(reports.resolve("key-0/capture/configuration-cache-report.html")));
    assertTrue(Files.exists(captureNote));
    assertFalse(Files.exists(reports.resolve("key-1")));
    assertTrue(Files.exists(reports.resolve("key-2/capture/configuration-cache-report.html")));
    assertTrue(Files.exists(reports.resolve("key-3/capture/configuration-cache-report.html")));
    assertTrue(Files.exists(reports.resolve("key-4/capture/configuration-cache-report.html")));
    assertTrue(Files.exists(other));
  }

  @Test
  void absentReportsNeedNoSetupOrCleanup() {
    var project = ProjectBuilder.builder().withProjectDir(root.toFile()).build();
    var task = GridGrindReportRetentionKt.registerGridGrindReportRetention(project).get();
    task.getActions().forEach(action -> action.execute(task));
    assertFalse(Files.exists(root.resolve("build/reports/configuration-cache")));
  }
}
