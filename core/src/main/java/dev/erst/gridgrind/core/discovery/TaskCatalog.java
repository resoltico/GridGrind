package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.List;
import java.util.Objects;

/** JSON-serializable task/intention catalog published for discovery. */
public record TaskCatalog(ProtocolVersion protocolVersion, List<TaskEntry> tasks) {
  public TaskCatalog {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    tasks = DiscoveryValidation.copyTaskEntries(tasks, "tasks");
  }
}
