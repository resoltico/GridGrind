package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import java.util.Objects;

/** Request-side configuration for execution-journal detail and rendering policy. */
@CatalogType(
    id = "ExecutionJournalInput",
    summary =
        "Explicit execution-journal settings. Use ExecutionJournalInput.defaults() for SUMMARY detail. SUMMARY returns compact, stable target summaries without timing telemetry, NORMAL returns expanded target summaries with timing telemetry, and VERBOSE also enables fine-grained event emission for CLI stderr rendering.",
    displayOrder = 11)
public record ExecutionJournalInput(
    @ProtocolField(optional = true)
        @JsonInclude(
            value = JsonInclude.Include.CUSTOM,
            valueFilter = ExecutionJournalInput.DefaultFilter.class)
        ExecutionJournalLevel level) {
  /** Returns the default journal input that keeps response telemetry compact and stable. */
  public static ExecutionJournalInput defaults() {
    return new ExecutionJournalInput(ExecutionJournalLevel.SUMMARY);
  }

  /** Reads one journal block while applying the documented omission default. */
  @JsonCreator
  static ExecutionJournalInput create(@JsonProperty("level") ExecutionJournalLevel level) {
    return new ExecutionJournalInput(level == null ? ExecutionJournalLevel.SUMMARY : level);
  }

  public ExecutionJournalInput {
    Objects.requireNonNull(level, "level must not be null");
  }

  /** Returns whether this input resolves to the product default journal behavior. */
  @JsonIgnore
  public boolean isDefault() {
    return level == ExecutionJournalLevel.SUMMARY;
  }

  /** Returns the required journal level after null/default normalization. */
  public static ExecutionJournalLevel effectiveLevel(ExecutionJournalInput journal) {
    return Objects.requireNonNull(journal, "journal must not be null").level();
  }

  /** Custom Jackson inclusion filter that omits the standard SUMMARY journal setting. */
  public static final class DefaultFilter {
    @Override
    public boolean equals(Object other) {
      return other == null
          || (other instanceof ExecutionJournalInput journal && journal.isDefault())
          || other == ExecutionJournalLevel.SUMMARY;
    }

    @Override
    public int hashCode() {
      return DefaultFilter.class.hashCode();
    }
  }
}
