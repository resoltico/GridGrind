package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import java.util.Objects;

/** Request-side configuration for execution-journal detail and rendering policy. */
@CatalogType(
    id = "ExecutionJournalSpec",
    summary =
        "Explicit execution-journal settings. Use ExecutionJournalSpec.defaults() for SUMMARY detail. SUMMARY returns compact, stable target summaries without timing telemetry, NORMAL returns expanded target summaries with timing telemetry, and VERBOSE also enables fine-grained event emission for CLI stderr rendering.",
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
