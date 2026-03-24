package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/** Tests for App process entry point and exit handler wiring. */
class AppTest {
  private App.CliFactory originalCliFactory;
  private App.ExitHandler originalExitHandler;
  private InputStream originalIn;
  private PrintStream originalOut;

  @BeforeEach
  void saveRuntimeHooks() {
    originalCliFactory = App.cliFactory;
    originalExitHandler = App.exitHandler;
    originalIn = System.in;
    originalOut = System.out;
  }

  @AfterEach
  void restoreRuntimeHooks() {
    App.cliFactory = originalCliFactory;
    App.exitHandler = originalExitHandler;
    System.setIn(originalIn);
    System.setOut(originalOut);
  }

  @Test
  void mainDelegatesToCliWithoutExitingOnSuccess() throws IOException {
    AtomicInteger observedExitCode = new AtomicInteger(-1);
    App.cliFactory =
        () ->
            (args, stdin, stdout) -> {
              assertArrayEquals(new String[] {"--request", "input.json"}, args);
              assertSame(System.in, stdin);
              assertSame(System.out, stdout);
              return 0;
            };
    App.exitHandler = observedExitCode::set;
    System.setIn(new ByteArrayInputStream(new byte[0]));
    System.setOut(new PrintStream(new ByteArrayOutputStream(), false, StandardCharsets.UTF_8));

    App.main(new String[] {"--request", "input.json"});

    assertEquals(-1, observedExitCode.get());
  }

  @Test
  void mainUsesExitHandlerForNonZeroExitCodes() throws IOException {
    AtomicInteger observedExitCode = new AtomicInteger(-1);
    App.cliFactory = () -> (args, stdin, stdout) -> 3;
    App.exitHandler = observedExitCode::set;
    System.setIn(new ByteArrayInputStream(new byte[0]));
    System.setOut(new PrintStream(new ByteArrayOutputStream(), false, StandardCharsets.UTF_8));

    App.main(new String[0]);

    assertEquals(3, observedExitCode.get());
  }
}
