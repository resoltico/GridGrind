package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;

/** JSON-serializable task/intention catalog emitted by the task-discovery CLI surface. */
public record TaskCatalog(GridGrindProtocolVersion protocolVersion, List<TaskEntry> tasks) {
  public TaskCatalog {
    protocolVersion =
        protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
    tasks = CatalogRecordValidation.copyTaskEntries(tasks, "tasks");
  }
}
