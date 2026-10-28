package dev.erst.gridgrind.fuzz;

import edu.berkeley.cs.jqf.fuzz.FuzzRunner;
import edu.berkeley.cs.jqf.fuzz.spi.ArgumentsGenerator;
import edu.berkeley.cs.jqf.fuzz.spi.ResultClassifier;
import edu.berkeley.cs.jqf.fuzz.spi.TrialExecutor;
import edu.berkeley.cs.jqf.instrument.tracing.SingleSnoop;
import edu.berkeley.cs.jqf.instrument.tracing.TraceLogger;
import java.io.File;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.Comparator;
import java.util.List;
import java.util.Random;
import java.util.concurrent.atomic.AtomicReference;

/** Standalone JUnit-free coverage driver; the same raw inputs still enter Harness.check. */
public final class CoverageFuzzing {
  private CoverageFuzzing() {}

  public static void main(String[] args) throws Throwable {
    if (args.length != 7)
      throw new IllegalArgumentException(
          "Expected harness, seeds, output, findings, seconds, trials and input bytes");
    Path expectedInstrumentation =
        Path.of(
            java.util.Objects.requireNonNull(
                System.getProperty("gridgrind.fuzz.instrumentation"),
                "Run active fuzzing through the Gradle fuzz task to select reviewed"
                    + " instrumentation"));
    for (Class<?> instrumenter :
        List.of(
            janala.instrument.SnoopInstructionTransformer.class,
            janala.instrument.SnoopInstructionClassAdapter.class,
            janala.instrument.SnoopInstructionMethodAdapter.class)) {
      Path loaded =
          Path.of(instrumenter.getProtectionDomain().getCodeSource().getLocation().toURI());
      if (!Files.isSameFile(expectedInstrumentation, loaded))
        throw new IllegalStateException(
            "Unexpected JQF instrumentation source for " + instrumenter.getName() + ": " + loaded);
    }
    System.out.println("Verified JQF instrumentation: " + expectedInstrumentation);
    Harness harness = Harness.fromKey(args[0]);
    Path seeds = Path.of(args[1]);
    Path output = Path.of(args[2]);
    Path findings = Path.of(args[3]);
    int seconds = Integer.parseInt(args[4]);
    long trials = Long.parseLong(args[5]);
    int maxInputBytes = Integer.parseInt(args[6]);
    System.clearProperty("jqfObservability");
    System.setProperty("jqf.ei.MAX_INPUT_SIZE", Integer.toString(maxInputBytes));
    System.setProperty("jqf.ei.EXIT_ON_CRASH", "true");
    System.setProperty("jqf.ei.LOG_CURRENT_INPUT", "false");
    System.setProperty("jqf.ei.LOG_ALL_INPUTS", "false");
    System.setProperty("jqf.ei.TIMEOUT", "10000");
    Files.createDirectories(findings);
    File[] seedFiles;
    try (var inputs = Files.list(seeds)) {
      seedFiles =
          inputs
              .filter(Files::isRegularFile)
              .sorted(Comparator.comparing(Path::toString))
              .map(Path::toFile)
              .toArray(File[]::new);
    }
    RawInputGuidance guidance =
        new RawInputGuidance(
            harness.key(),
            Duration.ofSeconds(seconds),
            trials,
            output.toFile(),
            findings.toFile(),
            seedFiles,
            maxInputBytes,
            new Random());
    AtomicReference<byte[]> inFlight = new AtomicReference<>();
    Runtime.getRuntime()
        .addShutdownHook(
            new Thread(
                () -> {
                  byte[] bytes = inFlight.get();
                  if (bytes != null) persist(findings, bytes, null);
                },
                "fuzz-input-evidence"));
    ArgumentsGenerator generator =
        input -> {
          try {
            byte[] bytes = input.readNBytes(maxInputBytes);
            if (input.read() != -1)
              throw new IllegalStateException("Guidance exceeded its input bound");
            return new Object[] {bytes};
          } catch (IOException exception) {
            throw new UncheckedIOException(exception);
          }
        };
    TrialExecutor executor =
        arguments -> {
          byte[] bytes = (byte[]) arguments[0];
          inFlight.set(bytes);
          harness.check(bytes);
          inFlight.set(null);
        };
    ResultClassifier classifier = error -> false;
    SingleSnoop.setCallbackGenerator(guidance::generateCallBack);
    TraceLogger.get().remove();
    SingleSnoop.startSnooping(Harness.class.getName() + "#check");
    try {
      new FuzzRunner(generator, executor, guidance, classifier, List.of(), null, true).run();
      if (guidance.firstFailure() != null) throw guidance.firstFailure();
      if (guidance.coverage() == 0)
        throw new IllegalStateException("No product coverage feedback was recorded");
      System.out.println(
          "JQF completed "
              + guidance.trials()
              + " trials; seeds executed="
              + guidance.seedTrialsExecuted()
              + "/"
              + guidance.seedCount()
              + "; product coverage probes="
              + guidance.coverage());
    } catch (Throwable error) {
      System.err.println(
          "JQF stopped after "
              + guidance.trials()
              + " trials; seeds executed="
              + guidance.seedTrialsExecuted()
              + "/"
              + guidance.seedCount()
              + "; product coverage probes="
              + guidance.coverage());
      Throwable primary = guidance.firstFailure() == null ? error : guidance.firstFailure();
      if (primary != error) primary.addSuppressed(error);
      persist(findings, inFlight.get(), primary);
      throw primary;
    } finally {
      inFlight.set(null);
      TraceLogger.get().remove();
    }
  }

  private static void persist(Path findings, byte[] bytes, Throwable primary) {
    if (bytes == null) return;
    try {
      Files.createDirectories(findings);
      try (var existing = Files.list(findings)) {
        if (existing.anyMatch(Files::isRegularFile)) return;
      }
      Files.write(findings.resolve("in-flight-input"), bytes);
      if (primary == null)
        Files.writeString(
            findings.resolve("in-flight-input.txt"),
            "Last in-flight input at process termination. Investigate timeout or cancellation; this"
                + " alone is not an invariant finding.\n");
    } catch (IOException exception) {
      if (primary != null) primary.addSuppressed(exception);
      else System.err.println("Could not preserve interrupted input: " + exception.getMessage());
    }
  }
}
