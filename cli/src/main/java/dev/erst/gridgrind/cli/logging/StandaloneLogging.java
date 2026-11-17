package dev.erst.gridgrind.cli.logging;

import dev.erst.gridgrind.cli.io.CliDiagnosticChannel;
import java.io.IOException;
import java.util.HashSet;
import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.core.LogEvent;
import org.apache.logging.log4j.core.LoggerContext;
import org.apache.logging.log4j.core.appender.AbstractAppender;
import org.apache.logging.log4j.core.config.Configurator;
import org.apache.logging.log4j.core.config.DefaultConfiguration;
import org.apache.logging.log4j.core.config.Property;
import org.apache.logging.log4j.status.StatusData;
import org.apache.logging.log4j.status.StatusListener;
import org.apache.logging.log4j.status.StatusLogger;

/** Explicit process-entrypoint policy; programmatic CLI and engine calls never install it. */
public final class StandaloneLogging implements AutoCloseable {
  private final NativeDiagnosticSink sink;
  private final StatusListener status;
  private final LoggerContext context;

  private StandaloneLogging(
      NativeDiagnosticSink sink, StatusListener status, LoggerContext context) {
    this.sink = sink;
    this.status = status;
    this.context = context;
  }

  /** Installs only the qualified synchronous metadata appender after JDK-only preflight. */
  @SuppressWarnings("PMD.CloseResource")
  // Successful context/listener ownership transfers to the returned closeable invocation.
  public static StandaloneLogging open(
      ClassLoader loader, String threshold, CliDiagnosticChannel channel) {
    NativeDiagnosticSink sink = new NativeDiagnosticSink(channel);
    StatusListener listener = new OwnedStatusListener(sink);
    StatusLogger.getLogger().registerListener(listener);
    boolean installed = false;
    try {
      var configuration = new DefaultConfiguration();
      configuration.setName("GridGrindStandaloneCli");
      var names = new HashSet<>(configuration.getAppenders().keySet());
      names.addAll(configuration.getRootLogger().getAppenders().keySet());
      for (String name : names) {
        configuration.getRootLogger().removeAppender(name);
        configuration.removeAppender(name);
      }
      configuration.getRootLogger().setLevel(Level.valueOf(threshold));
      var appender = new OwnedAppender(sink);
      appender.start();
      configuration.addAppender(appender);
      configuration.getRootLogger().addAppender(appender, Level.ALL, null);
      if (!configuration.getLoggers().isEmpty()
          || configuration.getAppenders().size() != 1
          || configuration.getRootLogger().getAppenders().size() != 1) {
        throw new IllegalStateException("Standalone logging must have exactly one owned appender");
      }
      LoggerContext context = Configurator.initialize(loader, configuration);
      if (context == null) throw new IllegalStateException("Standalone logging did not initialize");
      var logging = new StandaloneLogging(sink, listener, context);
      installed = true;
      return logging;
    } finally {
      if (!installed) StatusLogger.getLogger().removeListener(listener);
    }
  }

  @Override
  public void close() throws IOException {
    try {
      context.stop();
    } finally {
      StatusLogger.getLogger().removeListener(status);
    }
    sink.requireDelivery();
  }

  /** Sole synchronous appender; native messages and layouts never reach the channel. */
  private static final class OwnedAppender extends AbstractAppender {
    private final NativeDiagnosticSink sink;

    OwnedAppender(NativeDiagnosticSink sink) {
      super("GridGrindNativeDiagnostics", null, null, false, Property.EMPTY_ARRAY);
      this.sink = sink;
    }

    @Override
    public void append(LogEvent event) {
      sink.emit(
          NativeDiagnosticSink.component(event.getLoggerName()),
          event.getLevel(),
          event.getThrown());
    }
  }

  /** Uses the same bounded sink for provider initialization diagnostics. */
  private static final class OwnedStatusListener implements StatusListener {
    private final NativeDiagnosticSink sink;

    OwnedStatusListener(NativeDiagnosticSink sink) {
      this.sink = sink;
    }

    @Override
    public Level getStatusLevel() {
      return Level.WARN;
    }

    @Override
    public void log(StatusData data) {
      sink.emit("LOGGING", data.getLevel(), data.getThrowable());
    }

    @Override
    public void close() {
      /* No owned stream or asynchronous resource. */
    }
  }
}
