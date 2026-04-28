package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.executor.ExecutionJournalSink;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Objects;

/** Renders live verbose execution-journal events to stderr without affecting JSON responses. */
final class CliJournalWriter {
  ExecutionJournalSink sinkFor(WorkbookPlan request, OutputStream stderr) {
    Objects.requireNonNull(stderr, "stderr must not be null");
    if (request == null || request.journalLevel() != ExecutionJournalLevel.VERBOSE) {
      return ExecutionJournalSink.NOOP;
    }
    return event -> write(stderr, event);
  }

  private void write(OutputStream stderr, ExecutionJournal.Event event) {
    String line =
        "[gridgrind] "
            + event.timestamp()
            + " "
            + event.category()
            + event.stepId().map(stepId -> " stepId=" + stepId).orElse("")
            + event.stepIndex().map(stepIndex -> " stepIndex=" + stepIndex).orElse("")
            + " "
            + event.detail()
            + System.lineSeparator();
    try {
      stderr.write(line.getBytes(StandardCharsets.UTF_8));
      stderr.flush();
    } catch (IOException ignored) {
      // Best-effort journal rendering only; response JSON remains authoritative.
    }
  }
}
