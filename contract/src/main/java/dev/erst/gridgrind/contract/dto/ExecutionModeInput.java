package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonIgnore;
import java.util.Objects;

/** Optional top-level request settings that select low-memory execution families. */
public record ExecutionModeInput(ReadMode readMode, WriteMode writeMode) {
  private static final ReadMode DEFAULT_READ_MODE = ReadMode.FULL_XSSF;
  private static final WriteMode DEFAULT_WRITE_MODE = WriteMode.FULL_XSSF;

  /** Returns the default execution mode that keeps both reads and writes on full XSSF. */
  public static ExecutionModeInput defaults() {
    return new ExecutionModeInput(DEFAULT_READ_MODE, DEFAULT_WRITE_MODE);
  }

  public ExecutionModeInput {
    readMode = Objects.requireNonNullElse(readMode, DEFAULT_READ_MODE);
    writeMode = Objects.requireNonNullElse(writeMode, DEFAULT_WRITE_MODE);
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
