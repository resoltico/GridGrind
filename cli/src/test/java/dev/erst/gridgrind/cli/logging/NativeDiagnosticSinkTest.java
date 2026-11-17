package dev.erst.gridgrind.cli.logging;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.io.CliDiagnosticChannel;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import org.apache.logging.log4j.Level;
import org.junit.jupiter.api.Test;

/** Independent metadata, budget, transport-failure and frame-atomicity controls. */
class NativeDiagnosticSinkTest {
  @Test
  void customLevelsAndExceptionsNeverExposeMessagesAndShareOneBudget() throws Exception {
    var bytes = new ByteArrayOutputStream();
    var sink = new NativeDiagnosticSink(new CliDiagnosticChannel(bytes));
    var custom = Level.forName("synthetic-sensitive-custom-level", 100);
    var privateFailure = new IllegalArgumentException("private payload");
    for (int index = 0; index < 40; index++)
      sink.emit(index % 2 == 0 ? "POI" : "LOGGING", custom, privateFailure);
    sink.requireDelivery();
    var lines = bytes.toString(StandardCharsets.UTF_8).lines().toList();
    assertEquals(33, lines.size());
    assertEquals("{\"type\":\"NATIVE_DIAGNOSTIC_LIMIT\",\"limit\":32}", lines.getLast());
    assertTrue(lines.getFirst().contains("\"level\":\"FATAL\""));
    assertTrue(lines.getFirst().contains("java.lang.IllegalArgumentException"));
    assertFalse(bytes.toString(StandardCharsets.UTF_8).contains("private payload"));
    assertFalse(bytes.toString(StandardCharsets.UTF_8).contains("synthetic-sensitive"));
  }

  @Test
  void failedDiagnosticDeliveryIsVisibleWithoutRecursiveLoggingOrAnotherWrite() throws IOException {
    var attempts = new java.util.concurrent.atomic.AtomicInteger();
    var failure = new IOException("owned transport failure");
    try (var destination =
            new OutputStream() {
              @Override
              public void write(int value) throws IOException {
                attempts.incrementAndGet();
                throw failure;
              }
            };
        var channel = new CliDiagnosticChannel(destination)) {
      var sink = new NativeDiagnosticSink(channel);
      sink.emit("XML", Level.FATAL, null);
      sink.emit("LOGGING", Level.WARN, null);
      assertSame(failure, assertThrows(IOException.class, sink::requireDelivery));
      assertEquals(1, attempts.get());
    }
  }

  @Test
  void wholeFramesRemainAtomicAcrossNativeAndProgressWriters() throws Exception {
    var bytes = new ByteArrayOutputStream();
    var channel = new CliDiagnosticChannel(bytes);
    try (var threads = java.util.concurrent.Executors.newFixedThreadPool(2)) {
      var nativeWriter = threads.submit(() -> writeFrames(channel, "native"));
      var progressWriter = threads.submit(() -> writeFrames(channel, "progress"));
      nativeWriter.get();
      progressWriter.get();
    }
    var lines = bytes.toString(StandardCharsets.UTF_8).lines().toList();
    assertEquals(200, lines.size());
    assertTrue(
        lines.stream()
            .allMatch(line ->
                "{\"frame\":\"native\"}".equals(line) || "{\"frame\":\"progress\"}".equals(line)));
  }

  private static void writeFrames(CliDiagnosticChannel channel, String name) {
    for (int index = 0; index < 100; index++)
      try {
        channel.writeFrame(("{\"frame\":\"" + name + "\"}").getBytes(StandardCharsets.UTF_8));
      } catch (IOException failure) {
        throw new java.io.UncheckedIOException(failure);
      }
  }
}
