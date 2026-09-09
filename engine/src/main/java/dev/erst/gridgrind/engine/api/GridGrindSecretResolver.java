package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.SecretReference;

/** Resolves one host-authorized secret reference for a single GridGrind execution. */
@FunctionalInterface
public interface GridGrindSecretResolver {
  /** Resolves one reference into a caller-owned mutable secret buffer. */
  char[] resolve(SecretReference reference);
}
