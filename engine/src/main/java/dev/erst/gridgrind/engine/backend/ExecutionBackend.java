package dev.erst.gridgrind.engine.backend;

import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.semantics.BackendCapabilities;

/** Executes an admitted plan using its backend-specific resource ownership. */
public interface ExecutionBackend<C> {
  /** Returns the single shared capability declaration that owns routing and validation policy. */
  BackendCapabilities.Mode capabilities();

  /** Executes the backend call with its explicit resource ownership. */
  WorkbookResult execute(C call);
}
