package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.ExecutionProgressEvent;
import java.util.Objects;

/** Public live sink for structured execution progress from verbose plan execution. */
@FunctionalInterface
public interface GridGrindProgressSink {
  /** No-op sink used when live progress rendering is disabled. */
  GridGrindProgressSink NOOP = event -> {};

  /** Emits one structured execution-progress event. */
  void emit(ExecutionProgressEvent event);

  /** Returns a sink that rejects null delegates up front. */
  static GridGrindProgressSink requireNonNull(GridGrindProgressSink sink) {
    return Objects.requireNonNull(sink, "sink must not be null");
  }
}
