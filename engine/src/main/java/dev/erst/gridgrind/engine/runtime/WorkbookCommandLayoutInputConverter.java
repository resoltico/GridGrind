package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionSpec;
import java.util.Optional;

/** Converts named-range, protection, pane, and print/presentation contract inputs. */
final class WorkbookCommandLayoutInputConverter {
  private WorkbookCommandLayoutInputConverter() {}

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

  static ResolvedPrintLayout toResolvedPrintLayout(PrintLayoutSpec printLayout) {
    return new ResolvedPrintLayout(
        printLayout.printArea(),
        printLayout.orientation(),
        printLayout.scaling(),
        printLayout.repeatingRows(),
        printLayout.repeatingColumns(),
        new HeaderFooterTextReport(
            WorkbookCommandSourceSupport.inlineText(printLayout.header().left(), "header left"),
            WorkbookCommandSourceSupport.inlineText(printLayout.header().center(), "header center"),
            WorkbookCommandSourceSupport.inlineText(printLayout.header().right(), "header right")),
        new HeaderFooterTextReport(
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().left(), "footer left"),
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().center(), "footer center"),
            WorkbookCommandSourceSupport.inlineText(printLayout.footer().right(), "footer right")),
        printLayout.setup());
  }
}
