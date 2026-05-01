package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonIgnore;
import java.util.Objects;

/** Request-side configuration for execution-journal detail and rendering policy. */
public record ExecutionJournalInput(ExecutionJournalLevel level) {
  /** Returns the default journal input that keeps normal execution telemetry enabled. */
  public static ExecutionJournalInput defaults() {
    return new ExecutionJournalInput(ExecutionJournalLevel.NORMAL);
  }

  public ExecutionJournalInput {
    Objects.requireNonNull(level, "level must not be null");
  }

  /** Returns whether this input resolves to the product default journal behavior. */
  @JsonIgnore
  public boolean isDefault() {
    return level == ExecutionJournalLevel.NORMAL;
  }

  /** Returns the required journal level after null/default normalization. */
  public static ExecutionJournalLevel effectiveLevel(ExecutionJournalInput journal) {
    return Objects.requireNonNull(journal, "journal must not be null").level();
  }
}
