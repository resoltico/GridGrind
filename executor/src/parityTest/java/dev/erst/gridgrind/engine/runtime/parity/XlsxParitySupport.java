package dev.erst.gridgrind.engine.runtime.parity;

import dev.erst.gridgrind.engine.runtime.ExecutionInputBindings;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.nio.file.Path;
import java.util.Objects;
import java.util.concurrent.Callable;

/** Shared exception-wrapping support for the parity ledger, corpus, and oracle harness. */
final class XlsxParitySupport {
  private static final Path MANAGED_TEMP_SEGMENT = Path.of(".gridgrind", "tmp");
  private static final java.util.Set<String> CORPUS_ROOT_CHILDREN =
      java.util.Set.of("corpus", "scenario-copies", "workbooks");

  private XlsxParitySupport() {}

  static ExecutionInputBindings bindings(Path executionRoot) {
    return new ExecutionInputBindings(executionRoot, managedTempRoot(executionRoot));
  }

  static WorkbookTempFileFactory tempFileFactory(Path executionRoot) {
    return WorkbookTempFileFactory.rooted(managedTempRoot(executionRoot));
  }

  static Path executionRootFor(Path anchoredPath) {
    Path normalized = anchoredPath.toAbsolutePath().normalize();
    Path parent = normalized.getParent();
    for (Path current = parent; current != null; current = current.getParent()) {
      Path name = current.getFileName();
      if (name != null && CORPUS_ROOT_CHILDREN.contains(name.toString())) {
        Path corpusRoot = current.getParent();
        return corpusRoot == null ? current : corpusRoot;
      }
    }
    return parent == null ? normalized : parent;
  }

  private static Path managedTempRoot(Path executionRoot) {
    return executionRoot.toAbsolutePath().normalize().resolve(MANAGED_TEMP_SEGMENT);
  }

  static <T> T call(String action, Callable<T> callable) {
    Objects.requireNonNull(action, "action must not be null");
    Objects.requireNonNull(callable, "callable must not be null");
    try {
      return callable.call();
    } catch (RuntimeException exception) {
      throw exception;
    } catch (Exception exception) {
      throw new XlsxParityException(action, exception);
    }
  }
}

/** Runtime failure wrapper used when parity harness work cannot complete successfully. */
final class XlsxParityException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  XlsxParityException(String action, Exception cause) {
    super(action, cause);
  }
}
