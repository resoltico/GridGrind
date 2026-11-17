package dev.erst.gridgrind.cli.logging;

import dev.erst.gridgrind.cli.io.CliDiagnosticChannel;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Objects;
import java.util.Optional;
import org.apache.logging.log4j.Level;
import org.jspecify.annotations.Nullable;

/** Bounded safe native metadata shared by the owned appender and status listener. */
final class NativeDiagnosticSink {
  private static final int MAX_EVENTS = 32;
  private final CliDiagnosticChannel channel;
  private int events;
  private final java.util.concurrent.locks.ReentrantLock eventLock =
      new java.util.concurrent.locks.ReentrantLock();
  private Optional<IOException> failure = Optional.empty();

  NativeDiagnosticSink(CliDiagnosticChannel channel) {
    this.channel = Objects.requireNonNull(channel);
  }

  void emit(String component, Level level, @Nullable Throwable thrown) {
    eventLock.lock();
    try {
      if (failure.isPresent() || events > MAX_EVENTS) return;
      String frame;
      boolean truncated = events == MAX_EVENTS;
      events++;
      if (truncated) {
        frame = "{\"type\":\"NATIVE_DIAGNOSTIC_LIMIT\",\"limit\":" + MAX_EVENTS + "}";
      } else {
        String exception = thrown == null ? "" : safeException(thrown.getClass().getName());
        frame =
            "{\"type\":\"NATIVE_DIAGNOSTIC\",\"component\":\""
                + component
                + "\",\"level\":\""
                + level.getStandardLevel().name()
                + "\""
                + (exception.isEmpty() ? "" : ",\"exceptionType\":\"" + exception + "\"")
                + "}";
      }
      try {
        channel.writeFrame(frame.getBytes(StandardCharsets.UTF_8));
      } catch (IOException outputFailure) {
        failure = Optional.of(outputFailure);
      }
    } finally {
      eventLock.unlock();
    }
  }

  void requireDelivery() throws IOException {
    eventLock.lock();
    try {
      if (failure.isPresent()) throw failure.orElseThrow();
      channel.flush();
    } finally {
      eventLock.unlock();
    }
  }

  static String component(@Nullable String logger) {
    if (logger == null) return "NATIVE";
    if (logger.startsWith("org.apache.poi.")) return "POI";
    if (logger.startsWith("org.apache.xmlbeans.")
        || logger.startsWith("org.xml.")
        || logger.startsWith("javax.xml.")
        || logger.startsWith("com.ctc.wstx.")) return "XML";
    if (logger.startsWith("org.apache.xml.security.") || logger.startsWith("org.bouncycastle."))
      return "SECURITY";
    return "NATIVE";
  }

  private static String safeException(String name) {
    return name.length() <= 192
            && name.matches("[A-Za-z_$][A-Za-z0-9_$]*(\\.[A-Za-z_$][A-Za-z0-9_$]*)*")
        ? name
        : "Throwable";
  }
}
