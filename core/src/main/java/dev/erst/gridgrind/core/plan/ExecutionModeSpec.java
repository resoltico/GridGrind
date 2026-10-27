package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** Optional request-side execution-mode selector for the three published runtime families. */
@Union(discriminator = "type", catalog = true, displayOrder = 40)
public sealed interface ExecutionModeSpec {
  /** Returns the default execution mode that keeps reads and writes on the standard XSSF path. */
  static ExecutionModeSpec defaults() {
    return new FullXssf();
  }

  /** Returns the standard full-memory read/write execution mode. */
  static ExecutionModeSpec fullXssf() {
    return new FullXssf();
  }

  /** Returns the low-memory summary-read execution mode. */
  static ExecutionModeSpec eventRead() {
    return new EventRead();
  }

  /** Returns the low-memory append-write execution mode. */
  static ExecutionModeSpec streamingWrite() {
    return new StreamingWrite();
  }

  /** Stable SCREAMING_SNAKE_CASE discriminator used in the public contract. */
  String modeType();

  /** Returns whether this input resolves to the product default execution mode. */
  default boolean isDefault() {
    return this instanceof FullXssf;
  }

  /** Standard full-memory read/write execution mode with no special restrictions. */
  @Variant(id = "FULL_XSSF")
  record FullXssf() implements ExecutionModeSpec {
    @Override
    public String modeType() {
      return GridGrindProtocolTypeNames.executionModeTypeName(FullXssf.class);
    }
  }

  /** Low-memory summary-read execution mode with inspection-only constraints. */
  @Variant(id = "EVENT_READ")
  record EventRead() implements ExecutionModeSpec {
    @Override
    public String modeType() {
      return GridGrindProtocolTypeNames.executionModeTypeName(EventRead.class);
    }
  }

  /** Low-memory append-write execution mode with streaming mutation constraints. */
  @Variant(id = "STREAMING_WRITE")
  record StreamingWrite() implements ExecutionModeSpec {
    @Override
    public String modeType() {
      return GridGrindProtocolTypeNames.executionModeTypeName(StreamingWrite.class);
    }
  }
}
