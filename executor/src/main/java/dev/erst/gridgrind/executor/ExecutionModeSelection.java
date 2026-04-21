package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import java.util.Objects;

/** Resolved execution-mode pair selected for one request. */
record ExecutionModeSelection(
    ExecutionModeInput.ReadMode readMode, ExecutionModeInput.WriteMode writeMode) {
  ExecutionModeSelection {
    Objects.requireNonNull(readMode, "readMode must not be null");
    Objects.requireNonNull(writeMode, "writeMode must not be null");
  }
}
