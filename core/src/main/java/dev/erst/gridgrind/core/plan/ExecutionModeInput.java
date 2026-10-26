package dev.erst.gridgrind.core.plan;

import com.fasterxml.jackson.annotation.JsonIgnore;
import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** Optional request-side execution-mode selector for the three published runtime families. */
@Union(discriminator = "type", catalog = true, displayOrder = 40)
public sealed interface ExecutionModeInput {
  /** Returns the default execution mode that keeps reads and writes on the standard XSSF path. */
  static ExecutionModeInput defaults() {
    return new FullXssf();
  }

  /** Returns the standard full-memory read/write execution mode. */
  static ExecutionModeInput fullXssf() {
    return new FullXssf();
  }

  /** Returns the low-memory summary-read execution mode. */
  static ExecutionModeInput eventRead() {
    return new EventRead();
  }

  /** Returns the low-memory append-write execution mode. */
  static ExecutionModeInput streamingWrite() {
    return new StreamingWrite();
  }

  /** Stable SCREAMING_SNAKE_CASE discriminator used in the public contract. */
  String modeType();

  /** Returns whether this input resolves to the product default execution mode. */
  @JsonIgnore
  default boolean isDefault() {
    return this instanceof FullXssf;
  }

  /** Standard full-memory read/write execution mode with no special restrictions. */
  @Variant(
      id = "FULL_XSSF",
      summary =
          "Standard full workbook read and write execution with no special low-memory restrictions.")
  record FullXssf() implements ExecutionModeInput {
    @Override
    public String modeType() {
      return GridGrindProtocolTypeNames.executionModeTypeName(FullXssf.class);
    }
  }

  /** Low-memory summary-read execution mode with inspection-only constraints. */
  @Variant(
      id = "EVENT_READ",
      summary =
          "GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY only, with execution.calculation.strategy=DO_NOT_CALCULATE and markRecalculateOnOpen=false.")
  record EventRead() implements ExecutionModeInput {
    @Override
    public String modeType() {
      return GridGrindProtocolTypeNames.executionModeTypeName(EventRead.class);
    }
  }

  /** Low-memory append-write execution mode with streaming mutation constraints. */
  @Variant(
      id = "STREAMING_WRITE",
      summary =
          "source.type must be NEW; mutation actions limited to ENSURE_SHEET and APPEND_ROW; execution.calculation must keep strategy=DO_NOT_CALCULATE and may set markRecalculateOnOpen=true.")
  record StreamingWrite() implements ExecutionModeInput {
    @Override
    public String modeType() {
      return GridGrindProtocolTypeNames.executionModeTypeName(StreamingWrite.class);
    }
  }
}
