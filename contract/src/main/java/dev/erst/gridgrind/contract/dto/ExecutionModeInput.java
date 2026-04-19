package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonIgnore;
import java.util.Objects;

/** Optional top-level request settings that select low-memory execution families. */
public record ExecutionModeInput(ReadMode readMode, WriteMode writeMode) {
  public ExecutionModeInput {
    readMode = Objects.requireNonNullElse(readMode, ReadMode.FULL_XSSF);
    writeMode = Objects.requireNonNullElse(writeMode, WriteMode.FULL_XSSF);
  }

  /** Returns whether this execution-mode object leaves both reads and writes on full XSSF. */
  @JsonIgnore
  public boolean isDefault() {
    return readMode == ReadMode.FULL_XSSF && writeMode == WriteMode.FULL_XSSF;
  }

  /** Workbook read execution family. */
  public enum ReadMode {
    FULL_XSSF,
    EVENT_READ
  }

  /** Workbook write execution family. */
  public enum WriteMode {
    FULL_XSSF,
    STREAMING_WRITE
  }
}
