package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import java.util.Objects;

/** Request-side configuration for execution-journal detail and rendering policy. */
@CatalogType(
    id = "ExecutionJournalSpec",
    summary = dev.erst.gridgrind.core.catalog.GridGrindContractText.EXECUTION_JOURNAL_SUMMARY,
    displayOrder = 11)
public record ExecutionJournalSpec(
    @ProtocolField(optional = true) @WireDefault(value = "SUMMARY", omit = true)
        ExecutionJournalLevel level) {
  /** Returns the default journal input that keeps response telemetry compact and stable. */
  public static ExecutionJournalSpec defaults() {
    return new ExecutionJournalSpec(ExecutionJournalLevel.SUMMARY);
  }

  public ExecutionJournalSpec {
    Objects.requireNonNull(level, "level must not be null");
  }

  /** Returns whether this input resolves to the product default journal behavior. */
  public boolean isDefault() {
    return level == ExecutionJournalLevel.SUMMARY;
  }

  /** Returns the required journal level after null/default normalization. */
  public static ExecutionJournalLevel effectiveLevel(ExecutionJournalSpec journal) {
    return Objects.requireNonNull(journal, "journal must not be null").level();
  }
}
