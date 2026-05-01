package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.BooleanSupplier;
import org.junit.jupiter.api.Test;

/** Focused coverage tests for CLI stdin interactivity detection. */
class StandardInputInteractivityTest {
  @Test
  void currentProcessFactoryReturnsAnInstance() {
    StandardInputInteractivity interactivity = StandardInputInteractivity.currentProcess();

    assertNotNull(interactivity);
    interactivity.getAsBoolean();
  }

  @Test
  void neverFactoryReturnsSingletonInstance() {
    assertSame(StandardInputInteractivity.never(), StandardInputInteractivity.never());
  }

  @Test
  void reportsInteractiveWhenConsoleProbeSucceedsWithoutCallingNativeProbe() {
    AtomicInteger nativeProbeCalls = new AtomicInteger();
    StandardInputInteractivity interactivity =
        new StandardInputInteractivity(
            () -> true,
            () -> {
              nativeProbeCalls.incrementAndGet();
              return false;
            });

    assertTrue(interactivity.getAsBoolean());
    assertEquals(0, nativeProbeCalls.get());
  }

  @Test
  void reportsInteractiveWhenNativeProbeSucceeds() {
    StandardInputInteractivity interactivity =
        new StandardInputInteractivity(() -> false, () -> true);

    assertTrue(interactivity.getAsBoolean());
  }

  @Test
  void reportsNonInteractiveWhenBothProbesFail() {
    StandardInputInteractivity interactivity =
        new StandardInputInteractivity(() -> false, () -> false);

    assertFalse(interactivity.getAsBoolean());
  }

  @Test
  void consoleHelperSkipsTerminalProbeWhenNoConsoleIsPresent() {
    AtomicBoolean terminalProbeCalled = new AtomicBoolean();

    boolean interactive =
        StandardInputInteractivity.consoleIsInteractive(
            false,
            () -> {
              terminalProbeCalled.set(true);
              return true;
            });

    assertFalse(interactive);
    assertFalse(terminalProbeCalled.get());
  }

  @Test
  void consoleHelperReturnsTerminalProbeValueWhenConsoleIsPresent() {
    assertTrue(StandardInputInteractivity.consoleIsInteractive(true, () -> true));
    assertFalse(StandardInputInteractivity.consoleIsInteractive(true, () -> false));
  }

  @Test
  void lazyBooleanSupplierBuildsDelegateOnceAndReusesIt() {
    AtomicInteger factoryCalls = new AtomicInteger();
    AtomicInteger delegateCalls = new AtomicInteger();
    BooleanSupplier supplier =
        new LazyBooleanSupplier(
            () -> {
              factoryCalls.incrementAndGet();
              return () -> {
                delegateCalls.incrementAndGet();
                return true;
              };
            });

    assertTrue(supplier.getAsBoolean());
    assertTrue(supplier.getAsBoolean());
    assertEquals(1, factoryCalls.get());
    assertEquals(2, delegateCalls.get());
  }

  @Test
  void productionConstructorUsesInteractiveDetectorWithoutReadingProvidedInput() throws Exception {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    try (InputStream blockingStdin =
        new InputStream() {
          @Override
          public int read() {
            throw new AssertionError("interactive no-arg help must not read stdin");
          }

          @Override
          public int read(byte[] b, int off, int len) {
            throw new AssertionError("interactive no-arg help must not read stdin");
          }
        }) {
      int exitCode =
          new GridGrindCli(
                  (ignoredRequest, ignoredBindings, ignoredSink) -> {
                    throw new AssertionError("interactive no-arg help must not execute a request");
                  },
                  new CliRequestReader(),
                  new CliResponseWriter(),
                  new CliJournalWriter(),
                  () -> true)
              .run(new String[0], blockingStdin, stdout);

      assertEquals(0, exitCode);
      assertTrue(stdout.toString(StandardCharsets.UTF_8).contains("Usage:"));
    }
  }
}
