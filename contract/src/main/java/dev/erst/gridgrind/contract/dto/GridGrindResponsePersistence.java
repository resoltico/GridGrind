package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Persistence outcome variants reused by successful GridGrind responses. */
public interface GridGrindResponsePersistence {
  /** Reports whether the workbook was persisted during successful execution. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = PersistenceOutcome.NotSaved.class, name = "NONE"),
    @JsonSubTypes.Type(value = PersistenceOutcome.SavedAs.class, name = "SAVE_AS"),
    @JsonSubTypes.Type(value = PersistenceOutcome.Overwritten.class, name = "OVERWRITE")
  })
  sealed interface PersistenceOutcome
      permits PersistenceOutcome.NotSaved,
          PersistenceOutcome.SavedAs,
          PersistenceOutcome.Overwritten {

    /** Workbook remained in memory only and was not written to disk. */
    record NotSaved() implements PersistenceOutcome {}

    /**
     * Workbook was written to the path supplied in the SAVE_AS persistence field.
     *
     * <p>{@code requestedPath} is the literal string from the request. {@code executionPath} is the
     * absolute normalized path where the file was actually written. They differ when the request
     * supplies a relative path or a path with {@code ..} segments.
     */
    record SavedAs(String requestedPath, String executionPath) implements PersistenceOutcome {
      public SavedAs {
        Objects.requireNonNull(requestedPath, "requestedPath must not be null");
        Objects.requireNonNull(executionPath, "executionPath must not be null");
        if (requestedPath.isBlank()) {
          throw new IllegalArgumentException("requestedPath must not be blank");
        }
        if (executionPath.isBlank()) {
          throw new IllegalArgumentException("executionPath must not be blank");
        }
      }
    }

    /**
     * Workbook was saved by overwriting the opened source workbook path.
     *
     * <p>{@code sourcePath} is the path string from the EXISTING source as supplied in the request.
     * {@code executionPath} is the absolute normalized path where the file was written.
     */
    record Overwritten(String sourcePath, String executionPath) implements PersistenceOutcome {
      public Overwritten {
        Objects.requireNonNull(sourcePath, "sourcePath must not be null");
        Objects.requireNonNull(executionPath, "executionPath must not be null");
        if (sourcePath.isBlank()) {
          throw new IllegalArgumentException("sourcePath must not be blank");
        }
        if (executionPath.isBlank()) {
          throw new IllegalArgumentException("executionPath must not be blank");
        }
      }
    }
  }
}
