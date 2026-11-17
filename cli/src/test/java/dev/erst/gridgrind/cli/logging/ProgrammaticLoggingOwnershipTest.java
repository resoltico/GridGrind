package dev.erst.gridgrind.cli.logging;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.App;
import dev.erst.gridgrind.cli.command.GridGrindCli;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import org.apache.logging.log4j.core.LoggerContext;
import org.apache.logging.log4j.status.StatusLogger;
import org.junit.jupiter.api.Test;

/** Programmatic construction and invocation retain the host's existing provider and streams. */
class ProgrammaticLoggingOwnershipTest {
  @Test
  @SuppressWarnings("PMD.CloseResource") // System streams are borrowed host-owned resources.
  void programmaticEntryPointsPreserveTheHostConfigurationAndStatusListeners() throws Exception {
    var context = LoggerContext.getContext(false);
    var configuration = context.getConfiguration();
    var listeners = new java.util.ArrayList<org.apache.logging.log4j.status.StatusListener>();
    StatusLogger.getLogger().getListeners().forEach(listeners::add);
    var stdout = System.out;
    var stderr = System.err;
    assertNotNull(new App());
    assertEquals(
        0,
        new GridGrindCli()
            .run(
                new String[] {"version"},
                InputStream.nullInputStream(),
                new ByteArrayOutputStream(),
                new ByteArrayOutputStream()));
    assertSame(context, LoggerContext.getContext(false));
    assertSame(configuration, context.getConfiguration());
    var after = new java.util.ArrayList<org.apache.logging.log4j.status.StatusListener>();
    StatusLogger.getLogger().getListeners().forEach(after::add);
    assertEquals(listeners, after);
    assertSame(stdout, System.out);
    assertSame(stderr, System.err);
  }
}
