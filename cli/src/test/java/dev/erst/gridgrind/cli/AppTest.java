package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;

/** Tests for App process entry point and exit handler wiring. */
class AppTest {
  private static final String EMPTY_SUCCESS_REQUEST =
      """
      {
        "protocolVersion": "V1",
        "source": { "type": "NEW" },
        "persistence": { "type": "NONE" },
        "execution": {
          "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
          "journal": { "level": "NORMAL" },
          "calculation": {
            "strategy": { "type": "DO_NOT_CALCULATE" },
            "markRecalculateOnOpen": false
          }
        },
        "formulaEnvironment": {
          "externalWorkbooks": [],
          "missingWorkbookPolicy": "ERROR",
          "udfToolpacks": []
        },
        "steps": []
      }
      """;

  @Test
  void runDelegatesToCliAndDoesNotCallExitHandlerOnSuccess() throws IOException {
    AtomicInteger observedExitCode = new AtomicInteger(-1);
    AtomicReference<String> observedArgs = new AtomicReference<>();
    App app =
        new App(
            () ->
                (args, stdin, stdout, stderr) -> {
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
    App app = new App(() -> (args, stdin, stdout, stderr) -> 3, observedExitCode::set);

    app.run(new String[0]);

    assertEquals(3, observedExitCode.get());
  }

  @Test
  void defaultConstructorInitializesWithProductionDefaults() {
    assertNotNull(new App());
  }

  @Test
  void runWithProductionCliWiringConsumesStdinAndWritesResponse() throws IOException {
    byte[] jsonRequest = EMPTY_SUCCESS_REQUEST.getBytes(StandardCharsets.UTF_8);
    java.io.ByteArrayOutputStream capturedOut = new java.io.ByteArrayOutputStream();
    java.io.ByteArrayOutputStream capturedErr = new java.io.ByteArrayOutputStream();
    AtomicInteger observedExitCode = new AtomicInteger(-1);

    App app = new App(() -> new GridGrindCli()::run, observedExitCode::set);

    app.run(new String[0], new ByteArrayInputStream(jsonRequest), capturedOut, capturedErr);

    GridGrindResponse response = GridGrindJson.readResponse(capturedOut.toByteArray());

    assertEquals(-1, observedExitCode.get());
    assertInstanceOf(GridGrindResponse.Success.class, response);
    assertTrue(capturedErr.toString(StandardCharsets.UTF_8).isBlank());
  }

  @Test
  void mainConsumesSystemStreamsAndWritesResponse() throws IOException {
    byte[] jsonRequest = EMPTY_SUCCESS_REQUEST.getBytes(StandardCharsets.UTF_8);
    SystemStreams originalStreams = captureCurrentSystemStreams();
    ByteArrayOutputStream capturedOut = new ByteArrayOutputStream();
    ByteArrayOutputStream capturedErr = new ByteArrayOutputStream();
    try (PrintStream redirectedOut = new PrintStream(capturedOut, true, StandardCharsets.UTF_8);
        PrintStream redirectedErr = new PrintStream(capturedErr, true, StandardCharsets.UTF_8)) {
      System.setIn(new ByteArrayInputStream(jsonRequest));
      System.setOut(redirectedOut);
      System.setErr(redirectedErr);

      App.main(new String[0]);
    } finally {
      originalStreams.restore();
    }

    GridGrindResponse response = GridGrindJson.readResponse(capturedOut.toByteArray());
    assertInstanceOf(GridGrindResponse.Success.class, response);
    assertTrue(capturedErr.toString(StandardCharsets.UTF_8).isBlank());
  }

  private static SystemStreams captureCurrentSystemStreams() {
    return new SystemStreams(System.in, System.out, System.err);
  }

  /** Preserves and restores the mutable JVM-wide process streams during App entry-point tests. */
  private record SystemStreams(java.io.InputStream stdin, PrintStream stdout, PrintStream stderr) {
    void restore() {
      System.setIn(stdin);
      System.setOut(stdout);
      System.setErr(stderr);
    }
  }
}
