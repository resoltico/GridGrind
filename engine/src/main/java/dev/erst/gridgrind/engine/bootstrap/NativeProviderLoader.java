package dev.erst.gridgrind.engine.bootstrap;

import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.engine.api.GridGrindPayloadOpener;
import dev.erst.gridgrind.engine.api.GridGrindProgressSink;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestExecutor;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import dev.erst.gridgrind.engine.api.GridGrindSecretResolver;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Set;
import java.util.stream.Collectors;

/** One synchronous native implementation realm; it owns no stream beyond a class read. */
final class NativeProviderLoader extends ClassLoader {
  private static final String PREFIX = "META-INF/gridgrind/native-provider/";
  private static final Set<String> SHARED =
      Set.of(
          GridGrindRequestExecutor.class.getName(),
          GridGrindRequestDoctor.class.getName(),
          GridGrindRequestInputs.class.getName(),
          GridGrindPayloadOpener.class.getName(),
          GridGrindSecretResolver.class.getName(),
          GridGrindProgressSink.class.getName(),
          GridGrindEngine.class.getName());
  private final Set<String> required;

  NativeProviderLoader(ClassLoader resources) {
    super(java.util.Objects.requireNonNull(resources));
    try (var input =
        java.util.Objects.requireNonNull(getParent())
            .getResourceAsStream(PREFIX + "required-classes.list")) {
      if (input == null)
        throw new IllegalStateException("The native provider repair capsule is missing");
      required =
          new String(input.readAllBytes(), StandardCharsets.UTF_8)
              .lines()
              .filter(line -> !line.isBlank())
              .collect(Collectors.toUnmodifiableSet());
      if (required.isEmpty())
        throw new IllegalStateException("The native provider repair capsule is empty");
    } catch (IOException failure) {
      throw new IllegalStateException("Cannot read the native provider repair capsule", failure);
    }
  }

  private static boolean shared(String name) {
    return SHARED.stream().anyMatch(type -> name.equals(type) || name.startsWith(type + "$"));
  }

  private static boolean owned(String name) {
    return (name.startsWith("dev.erst.gridgrind.engine.")
            && !name.startsWith("dev.erst.gridgrind.engine.bootstrap.")
            && !shared(name))
        || name.startsWith("org.apache.poi.")
        || name.startsWith("org.apache.xmlbeans.")
        || name.startsWith("org.openxmlformats.schemas.")
        || name.startsWith("com.microsoft.schemas.")
        || name.startsWith("org.etsi.uri.")
        || name.startsWith("org.w3.x2000.")
        || name.startsWith("org.apache.xml.security.")
        || name.startsWith("org.apache.jcp.xml.dsig.");
  }

  @Override
  @SuppressWarnings("PMD.AvoidSynchronizedStatement")
  // ClassLoader's per-name loading lock coordinates with the inherited loading protocol.
  protected Class<?> loadClass(String name, boolean resolve) throws ClassNotFoundException {
    if (!owned(name)) return super.loadClass(name, resolve);
    synchronized (getClassLoadingLock(name)) {
      Class<?> loaded = findLoadedClass(name);
      if (loaded == null) {
        try (var input = classResource(name)) {
          byte[] bytes = input.readAllBytes();
          loaded = defineClass(name, bytes, 0, bytes.length);
        } catch (IOException failure) {
          throw new ClassNotFoundException(name, failure);
        }
      }
      if (resolve) resolveClass(loaded);
      return loaded;
    }
  }

  private java.io.InputStream classResource(String name) throws ClassNotFoundException {
    String path = name.replace('.', '/') + ".class";
    var parent = java.util.Objects.requireNonNull(getParent());
    var repaired = parent.getResourceAsStream(PREFIX + "classes/" + path);
    if (repaired != null) return repaired;
    if (required.contains(name))
      throw new ClassNotFoundException("Required native provider repair is missing: " + name);
    var original = parent.getResourceAsStream(path);
    if (original == null) throw new ClassNotFoundException(name);
    return original;
  }
}
