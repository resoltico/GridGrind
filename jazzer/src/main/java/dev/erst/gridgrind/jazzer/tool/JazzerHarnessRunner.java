package dev.erst.gridgrind.jazzer.tool;

import java.io.PrintWriter;
import java.util.Objects;
import org.junit.platform.engine.TestExecutionResult;
import org.junit.platform.engine.discovery.DiscoverySelectors;
import org.junit.platform.launcher.Launcher;
import org.junit.platform.launcher.LauncherDiscoveryRequest;
import org.junit.platform.launcher.TestExecutionListener;
import org.junit.platform.launcher.TestIdentifier;
import org.junit.platform.launcher.TestPlan;
import org.junit.platform.launcher.TagFilter;
import org.junit.platform.launcher.core.LauncherDiscoveryRequestBuilder;
import org.junit.platform.launcher.core.LauncherFactory;
import org.junit.platform.launcher.listeners.SummaryGeneratingListener;
import org.junit.platform.launcher.listeners.TestExecutionSummary;

/** Launches one Jazzer harness class through the JUnit Platform outside Gradle's Test task. */
public final class JazzerHarnessRunner {
  private static final String PULSE_PREFIX = "[JAZZER-PULSE] ";

  private JazzerHarnessRunner() {}

  /** Runs the requested Jazzer harness class and exits non-zero on any failure or misconfiguration. */
  public static void main(String[] args) {
    try (PrintWriter outputWriter = new PrintWriter(System.out, true);
        PrintWriter errorWriter = new PrintWriter(System.err, true)) {
      System.exit(run(parseClassName(args), outputWriter, errorWriter));
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
    return run(className, new PrintWriter(System.out, true), errorWriter);
  }

  /**
   * Executes one Jazzer harness class through the JUnit Platform and returns a process-style exit
   * code.
   */
  static int run(String className, PrintWriter outputWriter, PrintWriter errorWriter) {
    Objects.requireNonNull(className, "className must not be null");
    Objects.requireNonNull(outputWriter, "outputWriter must not be null");
    Objects.requireNonNull(errorWriter, "errorWriter must not be null");
    TestExecutionSummary summary = execute(className, outputWriter);
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

  private static TestExecutionSummary execute(String className, PrintWriter outputWriter) {
    LauncherDiscoveryRequest discoveryRequest =
        LauncherDiscoveryRequestBuilder.request()
            .selectors(DiscoverySelectors.selectClass(className))
            .filters(TagFilter.includeTags("jazzer"))
            .build();
    SummaryGeneratingListener listener = new SummaryGeneratingListener();
    Launcher launcher = LauncherFactory.create();
    launcher.registerTestExecutionListeners(listener, new PulseListener(className, outputWriter));
    launcher.execute(discoveryRequest);
    return listener.getSummary();
  }

  /** Emits concise per-harness progress pulses during standalone Jazzer launcher execution. */
  private static final class PulseListener implements TestExecutionListener {
    private final String className;
    private final PrintWriter outputWriter;
    private long totalTests;
    private long completedTests;

    private PulseListener(String className, PrintWriter outputWriter) {
      this.className = Objects.requireNonNull(className, "className must not be null");
      this.outputWriter = Objects.requireNonNull(outputWriter, "outputWriter must not be null");
    }

    @Override
    public void testPlanExecutionStarted(TestPlan testPlan) {
      totalTests = testPlan.countTestIdentifiers(TestIdentifier::isTest);
      outputWriter.println(
          PULSE_PREFIX + "harness-class=" + className + " phase=plan total-tests=" + totalTests);
    }

    @Override
    public void executionFinished(
        TestIdentifier testIdentifier, TestExecutionResult testExecutionResult) {
      if (!testIdentifier.isTest()) {
        return;
      }
      completedTests += 1;
      outputWriter.println(
          PULSE_PREFIX
              + "harness-class="
              + className
              + " phase=test-complete completed="
              + completedTests
              + "/"
              + totalTests
              + " status="
              + testExecutionResult.getStatus()
              + " test="
              + normalizedDisplayName(testIdentifier));
    }

    private String normalizedDisplayName(TestIdentifier testIdentifier) {
      return testIdentifier.getDisplayName().replaceAll("\\s+", " ").trim();
    }
  }
}
