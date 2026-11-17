package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.ExecutionStrategy;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactIo;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.io.IOException;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Retains prepared native resources until a workflow takes their exclusive ownership. */
final class PreparedWorkbookSource implements AutoCloseable {
  private final ExecutionStrategy strategy;
  private Optional<dev.erst.gridgrind.core.model.value.ExcelDateSystem> dateSystem =
      Optional.empty();
  private @Nullable XssfWorkbook xssf;
  private WorkbookArtifactIo.@Nullable MaterializedWorkbook materialized;

  PreparedWorkbookSource(ExecutionStrategy strategy) {
    this.strategy = Objects.requireNonNull(strategy);
  }

  ExecutionStrategy strategy() {
    return strategy;
  }

  void retain(XssfWorkbook workbook) {
    if (xssf != null || materialized != null) {
      throw new IllegalStateException("Source resource already prepared");
    }
    xssf = Objects.requireNonNull(workbook);
    dateSystem =
        Optional.of(
            workbook.xssfWorkbook().isDate1904()
                ? dev.erst.gridgrind.core.model.value.ExcelDateSystem.MAC_1904
                : dev.erst.gridgrind.core.model.value.ExcelDateSystem.WINDOWS_1900);
  }

  void retain(WorkbookArtifactIo.MaterializedWorkbook workbook) {
    if (xssf != null || materialized != null) {
      throw new IllegalStateException("Source resource already prepared");
    }
    materialized = Objects.requireNonNull(workbook);
  }

  Optional<dev.erst.gridgrind.core.model.value.ExcelDateSystem> dateSystem() {
    return dateSystem;
  }

  void newWorkbookCalendar() {
    dateSystem = Optional.of(dev.erst.gridgrind.core.model.value.ExcelDateSystem.WINDOWS_1900);
  }

  Optional<XssfWorkbook> takeXssf() {
    XssfWorkbook value = xssf;
    xssf = null;
    return Optional.ofNullable(value);
  }

  Optional<WorkbookArtifactIo.MaterializedWorkbook> takeMaterialized() {
    var value = materialized;
    materialized = null;
    return Optional.ofNullable(value);
  }

  @Override
  public void close() throws IOException {
    // Clearing before close makes the transfer/release obligation single-use even on failure.
    var full = takeXssf();
    var plain = takeMaterialized();
    IOException failure = null;
    if (full.isPresent()) {
      try {
        full.orElseThrow().close();
      } catch (IOException exception) {
        failure = exception;
      }
    }
    if (plain.isPresent()) {
      try {
        plain.orElseThrow().close();
      } catch (IOException exception) {
        if (failure == null) failure = exception;
        else failure.addSuppressed(exception);
      }
    }
    if (failure != null) throw failure;
  }
}
