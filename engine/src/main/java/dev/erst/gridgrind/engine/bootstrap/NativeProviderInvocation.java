package dev.erst.gridgrind.engine.bootstrap;

import dev.erst.gridgrind.core.semantics.ExecutionConventions;
import java.lang.reflect.InvocationTargetException;

/** Constructs a fresh native port for one public call, with no retained realm or worker. */
public final class NativeProviderInvocation {
  private NativeProviderInvocation() {}

  /** Creates the requested published port in a fresh realm using this artifact's own resources. */
  @SuppressWarnings({
    "PMD.UseProperClassLoader",
    "PMD.CompareObjectsWithEquals",
    "PMD.PreserveStackTrace"
  })
  // Artifact-loader identity is the boundary; unwrap the real cause without replacing fatal
  // identity.
  public static <T> T create(Class<T> port) {
    var loader =
        new NativeProviderLoader(
            java.util.Objects.requireNonNull(NativeProviderInvocation.class.getClassLoader()));
    try {
      Class<?> implementation =
          Class.forName("dev.erst.gridgrind.engine.runtime.NativeEnginePort", true, loader);
      if (implementation.getClassLoader() != loader)
        throw new IllegalStateException("Native entry point escaped its owned provider realm");
      return implementation
          .asSubclass(port)
          .getConstructor(ExecutionConventions.class)
          .newInstance(ExecutionConventions.STANDARD);
    } catch (InvocationTargetException wrapper) {
      if (wrapper.getCause() instanceof Error error) throw error;
      if (wrapper.getCause() instanceof RuntimeException runtime) throw runtime;
      throw new IllegalStateException("Native provider construction failed", wrapper.getCause());
    } catch (ReflectiveOperationException failure) {
      throw new IllegalStateException("Native provider entry point is incompatible", failure);
    }
  }
}
