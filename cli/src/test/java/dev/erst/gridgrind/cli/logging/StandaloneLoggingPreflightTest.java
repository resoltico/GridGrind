package dev.erst.gridgrind.cli.logging;

import static org.junit.jupiter.api.Assertions.*;

import java.net.URLClassLoader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Rejects routing/provider policy before any provider class can initialize. */
class StandaloneLoggingPreflightTest {
  @TempDir Path root;

  @Test
  void acceptsOnlyConsistentStandardVerbosityAliases() {
    assertEquals(
        "DEBUG", StandaloneLoggingPreflight.level(Map.of("log4j2.level", " debug ")).orElseThrow());
    assertEquals(
        "TRACE",
        StandaloneLoggingPreflight.level(
                Map.of("LOG4J_LEVEL", "TRACE", "org.apache.logging.log4j.level", "trace"))
            .orElseThrow());
    assertThrows(
        StandaloneLoggingRefusal.class,
        () ->
            StandaloneLoggingPreflight.level(
                Map.of("log4j2.level", "WARN", "LOG4J_LEVEL", "DEBUG")));
    for (String value : new String[] {"ERROR", "FATAL", "OFF", "sensitive-custom-level"})
      assertThrows(
          StandaloneLoggingRefusal.class,
          () -> StandaloneLoggingPreflight.level(Map.of("log4j2.level", value)));
  }

  @Test
  void refusesCanonicalLegacyEnvironmentAndFuzzyReservedRoutingWithoutEchoingValues() {
    for (String key :
        new String[] {
          "log4j.loggerContextFactory",
          "LOG4J_CONFIGURATION_FILE",
          "Log4jContextSelector",
          "log4j2.statusLoggerLevel",
          "asyncLoggerRingBufferSize",
          "disableThreadContext",
          "log4j2.fuzzy.unqualified"
        }) {
      var failure =
          assertThrows(
              StandaloneLoggingRefusal.class,
              () -> StandaloneLoggingPreflight.level(Map.of(key, "private-value")));
      assertEquals(key, failure.setting());
      assertFalse(failure.getMessage().contains("private-value"));
    }
    assertTrue(StandaloneLoggingPreflight.level(Map.of("ordinary-setting", "value")).isEmpty());
  }

  @Test
  void refusesPackagedPropertySourcesWithoutLoadingTheirContents() throws Exception {
    Files.writeString(
        root.resolve("log4j2.system.properties"),
        "log4j.loggerContextFactory=must.never.initialize.Provider");
    try (var loader = new URLClassLoader(new java.net.URL[] {root.toUri().toURL()}, null)) {
      var failure =
          assertThrows(
              StandaloneLoggingRefusal.class, () -> StandaloneLoggingPreflight.validate(loader));
      assertEquals("log4j2.system.properties", failure.setting());
    }
  }

  @Test
  void boundsAndValidatesProviderDescriptorsWithoutClassLoading() throws Exception {
    Path descriptor = root.resolve("META-INF/services/org.apache.logging.log4j.spi.Provider");
    Files.createDirectories(descriptor.getParent());
    try (var loader = new URLClassLoader(new java.net.URL[] {root.toUri().toURL()}, null)) {
      Files.writeString(
          descriptor,
          "# allowed implementation\norg.apache.logging.log4j.core.impl.Log4jProvider\n");
      assertNotNull(StandaloneLoggingPreflight.validate(loader));
      for (String unsafe :
          new String[] {
            "# comment\nmust.never.initialize.Provider\n",
            "org.apache.logging.log4j.core.impl.Log4jProvider\nmust.never.initialize.Provider\n",
            "# comment\r\nmust.never.initialize.Provider\r\n"
          }) {
        Files.writeString(descriptor, unsafe);
        assertThrows(
            StandaloneLoggingRefusal.class, () -> StandaloneLoggingPreflight.validate(loader));
      }
      Files.writeString(descriptor, "must.never.initialize.Provider");
      assertThrows(
          StandaloneLoggingRefusal.class, () -> StandaloneLoggingPreflight.validate(loader));
      Files.write(descriptor, new byte[] {(byte) 0xff});
      assertThrows(
          StandaloneLoggingRefusal.class, () -> StandaloneLoggingPreflight.validate(loader));
      Files.writeString(descriptor, "#" + "x".repeat(8191));
      assertNotNull(StandaloneLoggingPreflight.validate(loader));
      Files.writeString(descriptor, "#" + "x".repeat(8192));
      assertThrows(
          StandaloneLoggingRefusal.class, () -> StandaloneLoggingPreflight.validate(loader));
    }
  }
}
