package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;

/** JSON-serializable task/intention catalog emitted by the task-discovery CLI surface. */
public record TaskCatalog(GridGrindProtocolVersion protocolVersion, List<TaskEntry> tasks) {
  public TaskCatalog {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    tasks = CatalogRecordValidation.copyTaskEntries(tasks, "tasks");
  }
}
