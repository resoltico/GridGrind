package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
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
    Objects.requireNonNull(readMode, "readMode must not be null");
    Objects.requireNonNull(writeMode, "writeMode must not be null");
  }

  /** Creates an execution mode that keeps reads on the default family. */
  public ExecutionModeInput(WriteMode writeMode) {
    this(DEFAULT_READ_MODE, writeMode);
  }

  @JsonCreator
  static ExecutionModeInput create(
      @JsonProperty("readMode") ReadMode readMode, @JsonProperty("writeMode") WriteMode writeMode) {
    return new ExecutionModeInput(
        readMode == null ? DEFAULT_READ_MODE : readMode,
        writeMode == null ? DEFAULT_WRITE_MODE : writeMode);
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
