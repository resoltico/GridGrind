package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonIgnore;
import java.util.Objects;

/** Request-side configuration for execution-journal detail and rendering policy. */
public record ExecutionJournalInput(ExecutionJournalLevel level) {
  public ExecutionJournalInput {
    level = level == null ? ExecutionJournalLevel.NORMAL : level;
  }

  /** Returns whether this input resolves to the product default journal behavior. */
  @JsonIgnore
  public boolean isDefault() {
    return level == ExecutionJournalLevel.NORMAL;
  }

  /** Returns the required journal level after null/default normalization. */
  public static ExecutionJournalLevel effectiveLevel(ExecutionJournalInput journal) {
    return Objects.requireNonNullElse(journal, new ExecutionJournalInput(null)).level();
  }
}
