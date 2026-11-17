package dev.erst.gridgrind.cli.logging;

import java.io.IOException;
import java.net.JarURLConnection;
import java.net.URL;
import java.nio.ByteBuffer;
import java.nio.charset.CodingErrorAction;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

/** JDK-only admission before any active logging-provider initialization. */
public final class StandaloneLoggingPreflight {
  private static final Set<String> LEVEL_KEYS =
      Set.of("log4jlevel", "log4j2level", "orgapachelogginglog4jlevel");
  private static final Set<String> LEVELS = Set.of("WARN", "INFO", "DEBUG", "TRACE", "ALL");
  private static final int MAX_DESCRIPTOR_BYTES = 8192;

  private StandaloneLoggingPreflight() {}

  /** Returns a qualified standard threshold; arbitrary provider and routing tuning refuses. */
  public static String validate(ClassLoader loader) throws IOException {
    Optional<String> properties = level(System.getProperties());
    Optional<String> environment = level(System.getenv());
    for (String role :
        List.of(
            "log4j2.system.properties",
            "log4j2.component.properties",
            "log4j2.simplelog.properties",
            "log4j2.StatusLogger.properties",
            "META-INF/log4j-provider.properties")) {
      if (loader.getResources(role).hasMoreElements()) throw new StandaloneLoggingRefusal(role);
    }
    Map<String, Set<String>> services =
        Map.of(
            "org.apache.logging.log4j.spi.Provider",
                Set.of("org.apache.logging.log4j.core.impl.Log4jProvider"),
            "org.apache.logging.log4j.util.PropertySource",
                Set.of(
                    "org.apache.logging.log4j.util.EnvironmentPropertySource",
                    "org.apache.logging.log4j.util.SystemPropertiesPropertySource"),
            "org.apache.logging.log4j.status.StatusListener", Set.of());
    for (var service : services.entrySet())
      inspectServices(loader, service.getKey(), service.getValue());
    return properties.or(() -> environment).orElse("WARN");
  }

  static Optional<String> level(Map<?, ?> settings) {
    Optional<String> selected = Optional.empty();
    for (var setting : settings.entrySet()) {
      String key = String.valueOf(setting.getKey());
      String normalized = key.replaceAll("[^A-Za-z0-9]", "").toLowerCase(Locale.ROOT);
      if (!reserved(normalized)) continue;
      // Report a fixed role when a supplied key itself is not a bounded safe identifier.
      String safeKey = key.matches("[A-Za-z0-9_.-]{1,128}") ? key : "logging setting";
      if (!LEVEL_KEYS.contains(normalized)) throw new StandaloneLoggingRefusal(safeKey);
      String threshold = String.valueOf(setting.getValue()).strip().toUpperCase(Locale.ROOT);
      if (!LEVELS.contains(threshold)
          || (selected.isPresent() && !selected.orElseThrow().equals(threshold))) {
        throw new StandaloneLoggingRefusal(safeKey);
      }
      selected = Optional.of(threshold);
    }
    return selected;
  }

  private static boolean reserved(String key) {
    return key.startsWith("log4j")
        || key.startsWith("orgapachelogginglog4j")
        || key.startsWith("asynclogger")
        || key.startsWith("disablethreadcontext")
        || key.startsWith("isthreadcontext");
  }

  private static void inspectServices(ClassLoader loader, String service, Set<String> allowed)
      throws IOException {
    String role = "META-INF/services/" + service;
    var resources = loader.getResources(role);
    while (resources.hasMoreElements()) {
      URL resource = resources.nextElement();
      requireLocal(resource, role);
      byte[] bytes;
      try (var input = resource.openStream()) {
        bytes = input.readNBytes(MAX_DESCRIPTOR_BYTES + 1);
      }
      if (bytes.length > MAX_DESCRIPTOR_BYTES) throw new StandaloneLoggingRefusal(role);
      String descriptor;
      try {
        descriptor =
            StandardCharsets.UTF_8
                .newDecoder()
                .onMalformedInput(CodingErrorAction.REPORT)
                .onUnmappableCharacter(CodingErrorAction.REPORT)
                .decode(ByteBuffer.wrap(bytes))
                .toString();
      } catch (java.nio.charset.CharacterCodingException malformed) {
        throw new StandaloneLoggingRefusal(role, malformed);
      }
      for (String line : descriptor.lines().toList()) {
        int comment = line.indexOf('#');
        String provider = (comment < 0 ? line : line.substring(0, comment)).trim();
        if (!provider.isEmpty() && !allowed.contains(provider))
          throw new StandaloneLoggingRefusal(role);
      }
    }
  }

  private static void requireLocal(URL resource, String role) throws IOException {
    if ("file".equals(resource.getProtocol())) return;
    if ("jar".equals(resource.getProtocol())
        && resource.openConnection() instanceof JarURLConnection jar
        && "file".equals(jar.getJarFileURL().getProtocol())) return;
    throw new StandaloneLoggingRefusal(role);
  }
}
