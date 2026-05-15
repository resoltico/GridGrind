package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;

/** Optional request-side execution-mode selector for the three published runtime families. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ExecutionModeInput.FullXssf.class, name = "FULL_XSSF"),
  @JsonSubTypes.Type(value = ExecutionModeInput.EventRead.class, name = "EVENT_READ"),
  @JsonSubTypes.Type(value = ExecutionModeInput.StreamingWrite.class, name = "STREAMING_WRITE")
})
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
  record FullXssf() implements ExecutionModeInput {
    @Override
    public String modeType() {
      return GridGrindProtocolTypeNames.executionModeTypeName(FullXssf.class);
    }
  }

  /** Low-memory summary-read execution mode with inspection-only constraints. */
  record EventRead() implements ExecutionModeInput {
    @Override
    public String modeType() {
      return GridGrindProtocolTypeNames.executionModeTypeName(EventRead.class);
    }
  }

  /** Low-memory append-write execution mode with streaming mutation constraints. */
  record StreamingWrite() implements ExecutionModeInput {
    @Override
    public String modeType() {
      return GridGrindProtocolTypeNames.executionModeTypeName(StreamingWrite.class);
    }
  }
}
