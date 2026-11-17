package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import dev.erst.gridgrind.core.plan.PreparedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import java.util.Optional;

/** Acquires separately granted secrets at the native mutation boundary. */
final class MutationSecrets {
  private MutationSecrets() {}

  static ResolvedMutationAction finish(
      PreparedMutationAction action, ExecutionInputBindings bindings) {
    return switch (action) {
      case ResolvedMutationAction resolved -> resolved;
      case WorkbookMutationAction.SetSheetProtection protection ->
          new ResolvedMutationAction.SetSheetProtection(
              protection.protection(), protection.passwordRef().map(bindings::resolveSecret));
      case WorkbookMutationAction.SetWorkbookProtection protection ->
          new ResolvedMutationAction.SetWorkbookProtection(
              toResolvedWorkbookProtection(protection.protection(), bindings));
    };
  }

  static ResolvedWorkbookProtection toResolvedWorkbookProtection(
      WorkbookProtectionSpec protection, ExecutionInputBindings bindings) {
    return new ResolvedWorkbookProtection(
        protection.structureLocked(),
        protection.windowsLocked(),
        protection.revisionsLocked(),
        resolveOptionalSecret(protection.workbookPasswordRef(), bindings),
        resolveOptionalSecret(protection.revisionsPasswordRef(), bindings));
  }

  private static Optional<String> resolveOptionalSecret(
      Optional<dev.erst.gridgrind.core.source.SecretReference> reference,
      ExecutionInputBindings bindings) {
    if (reference.isEmpty()) {
      return Optional.empty();
    }
    if (bindings == null) {
      throw new IllegalStateException(
          "secret-bearing workbook protection requires execution bindings");
    }
    return Optional.of(bindings.resolveSecret(reference.orElseThrow()));
  }
}
