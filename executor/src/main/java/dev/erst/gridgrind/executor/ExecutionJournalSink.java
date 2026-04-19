package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import java.util.Objects;

/** Live sink for verbose execution-journal events emitted during plan execution. */
@FunctionalInterface
public interface ExecutionJournalSink {
  /** No-op sink used when live journal rendering is disabled. */
  ExecutionJournalSink NOOP = event -> {};

  /** Emits one live execution-journal event. */
  void emit(ExecutionJournal.Event event);

  /** Returns a sink that rejects null delegates up front. */
  static ExecutionJournalSink requireNonNull(ExecutionJournalSink sink) {
    return Objects.requireNonNull(sink, "sink must not be null");
  }
}
