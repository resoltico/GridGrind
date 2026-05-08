package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import java.util.Objects;

/** Live sink for verbose execution-journal events emitted during plan execution. */
@FunctionalInterface
public interface GridGrindJournalSink {
  /** No-op sink used when live journal rendering is disabled. */
  GridGrindJournalSink NOOP = event -> {};

  /** Emits one live execution-journal event. */
  void emit(ExecutionJournal.Event event);

  /** Returns a sink that rejects null delegates up front. */
  static GridGrindJournalSink requireNonNull(GridGrindJournalSink sink) {
    return Objects.requireNonNull(sink, "sink must not be null");
  }
}
