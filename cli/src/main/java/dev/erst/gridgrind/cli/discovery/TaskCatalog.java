package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.List;
import java.util.Objects;

/** JSON-serializable task/intention catalog emitted by the task-discovery CLI surface. */
public record TaskCatalog(ProtocolVersion protocolVersion, List<TaskEntry> tasks) {
  public TaskCatalog {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    tasks = CliDiscoveryValidation.copyTaskEntries(tasks, "tasks");
  }
}
