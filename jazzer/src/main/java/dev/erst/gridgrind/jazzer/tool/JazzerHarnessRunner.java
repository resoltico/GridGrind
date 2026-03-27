package dev.erst.gridgrind.jazzer.tool;

import java.io.PrintWriter;
import java.util.Objects;
import org.junit.platform.engine.discovery.DiscoverySelectors;
import org.junit.platform.launcher.Launcher;
import org.junit.platform.launcher.LauncherDiscoveryRequest;
import org.junit.platform.launcher.TagFilter;
import org.junit.platform.launcher.core.LauncherDiscoveryRequestBuilder;
import org.junit.platform.launcher.core.LauncherFactory;
import org.junit.platform.launcher.listeners.SummaryGeneratingListener;
import org.junit.platform.launcher.listeners.TestExecutionSummary;

/** Launches one Jazzer harness class through the JUnit Platform outside Gradle's Test task. */
public final class JazzerHarnessRunner {
  private JazzerHarnessRunner() {}

  /** Runs the requested Jazzer harness class and exits non-zero on any failure or misconfiguration. */
  public static void main(String[] args) {
    try (PrintWriter errorWriter = new PrintWriter(System.err, true)) {
      System.exit(run(parseClassName(args), errorWriter));
    }
  }

  /** Parses the required `--class <fqcn>` argument pair for launcher-based Jazzer execution. */
  static String parseClassName(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    if (args.length != 2 || !"--class".equals(args[0])) {
      throw new IllegalArgumentException("Usage: JazzerHarnessRunner --class <fully-qualified-class>");
    }
    String className = Objects.requireNonNull(args[1], "className must not be null");
    if (className.isBlank()) {
      throw new IllegalArgumentException("className must not be blank");
    }
    return className;
  }

  /**
   * Executes one Jazzer harness class through the JUnit Platform and returns a process-style exit
   * code.
   */
  static int run(String className, PrintWriter errorWriter) {
    Objects.requireNonNull(className, "className must not be null");
    Objects.requireNonNull(errorWriter, "errorWriter must not be null");
    TestExecutionSummary summary = execute(className);
    if (summary.getTestsFoundCount() == 0) {
      errorWriter.println("No Jazzer tests were discovered for class: " + className);
      return 1;
    }
    if (summary.getTotalFailureCount() > 0) {
      summary.printFailuresTo(errorWriter);
      return 1;
    }
    return 0;
  }

  private static TestExecutionSummary execute(String className) {
    LauncherDiscoveryRequest discoveryRequest =
        LauncherDiscoveryRequestBuilder.request()
            .selectors(DiscoverySelectors.selectClass(className))
            .filters(TagFilter.includeTags("jazzer"))
            .build();
    SummaryGeneratingListener listener = new SummaryGeneratingListener();
    Launcher launcher = LauncherFactory.create();
    launcher.registerTestExecutionListeners(listener);
    launcher.execute(discoveryRequest);
    return listener.getSummary();
  }
}
