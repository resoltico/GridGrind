package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;

/** Tests for App process entry point and exit handler wiring. */
class AppTest {
  @Test
  void runDelegatesToCliAndDoesNotCallExitHandlerOnSuccess() throws IOException {
    AtomicInteger observedExitCode = new AtomicInteger(-1);
    AtomicReference<String> observedArgs = new AtomicReference<>();
    App app =
        new App(
            () ->
                (args, stdin, stdout) -> {
                  observedArgs.set(String.join(",", args));
                  return 0;
                },
            observedExitCode::set);

    app.run(new String[] {"--request", "input.json"});

    assertEquals("--request,input.json", observedArgs.get());
    assertEquals(-1, observedExitCode.get());
  }

  @Test
  void runCallsExitHandlerForNonZeroExitCodes() throws IOException {
    AtomicInteger observedExitCode = new AtomicInteger(-1);
    App app = new App(() -> (args, stdin, stdout) -> 3, observedExitCode::set);

    app.run(new String[0]);

    assertEquals(3, observedExitCode.get());
  }

  @Test
  void defaultConstructorInitializesWithProductionDefaults() {
    assertNotNull(new App());
  }

  @Test
  @SuppressWarnings("PMD.CloseResource")
  void mainMethodRunsEndToEndWithValidRequestOnStdin() throws IOException {
    byte[] jsonRequest =
        "{\"source\":{\"mode\":\"NEW\"},\"operations\":[],\"analysis\":{\"sheets\":[]}}"
            .getBytes(StandardCharsets.UTF_8);
    java.io.ByteArrayOutputStream capturedOut = new java.io.ByteArrayOutputStream();
    InputStream originalIn = System.in;
    PrintStream originalOut = System.out;
    try {
      System.setIn(new ByteArrayInputStream(jsonRequest));
      System.setOut(new PrintStream(capturedOut, false, StandardCharsets.UTF_8));
      App.main(new String[0]);
    } finally {
      System.setIn(originalIn);
      System.setOut(originalOut);
    }
    assertFalse(capturedOut.toString(StandardCharsets.UTF_8).isBlank());
  }
}
