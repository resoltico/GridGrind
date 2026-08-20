package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ExecutionProgressEvent;
import java.util.Objects;

/** Live sink for structured execution progress emitted during verbose plan execution. */
@FunctionalInterface
public interface ExecutionProgressSink {
  /** No-op sink used when live progress rendering is disabled. */
  ExecutionProgressSink NOOP = event -> {};

  /** Emits one structured execution-progress event. */
  void emit(ExecutionProgressEvent event);

  /** Returns a sink that rejects null delegates up front. */
  static ExecutionProgressSink requireNonNull(ExecutionProgressSink sink) {
    return Objects.requireNonNull(sink, "sink must not be null");
  }
}
