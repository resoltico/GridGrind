package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;
import java.util.Optional;

/** Persistence outcome variants returned on every workbook result. */
public interface WorkbookResultPersistence {
  /** Reports whether the workbook was persisted during execution. */
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
     * Workbook targeted the path supplied in the SAVE_AS persistence field.
     *
     * <p>{@code requestedPath} is the literal string from the request. {@code write} reports
     * whether the file was written and, when successful, the absolute normalized path where the
     * file was actually written.
     */
    record SavedAs(String requestedPath, WriteResult write) implements PersistenceOutcome {
      public SavedAs {
        Objects.requireNonNull(requestedPath, "requestedPath must not be null");
        Objects.requireNonNull(write, "write must not be null");
        if (requestedPath.isBlank()) {
          throw new IllegalArgumentException("requestedPath must not be blank");
        }
      }
    }

    /**
     * Workbook targeted the opened source workbook path for overwrite persistence.
     *
     * <p>{@code sourcePath} echoes the path string from an {@code EXISTING} source when that source
     * path is available. It is omitted when validation fails before any source workbook path
     * exists, such as an invalid {@code OVERWRITE} request paired with {@code source.type=NEW}.
     * {@code write} reports whether the target file was written and, when successful, the absolute
     * normalized path that was updated.
     */
    record Overwritten(
        @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> sourcePath, WriteResult write)
        implements PersistenceOutcome {
      public Overwritten {
        sourcePath = normalizeSourcePath(sourcePath);
        Objects.requireNonNull(write, "write must not be null");
      }

      /**
       * Creates one overwrite outcome that echoes a known EXISTING source path from the request.
       */
      public Overwritten(String sourcePath, WriteResult write) {
        this(Optional.of(sourcePath), write);
      }

      private static Optional<String> normalizeSourcePath(Optional<String> sourcePath) {
        Optional<String> normalized = Objects.requireNonNullElseGet(sourcePath, Optional::empty);
        if (normalized.isEmpty()) {
          return Optional.empty();
        }
        String path = normalized.orElseThrow();
        if (path.isBlank()) {
          throw new IllegalArgumentException("sourcePath must not be blank");
        }
        return Optional.of(path);
      }
    }
  }

  /** Reports whether the targeted workbook file was actually written. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WriteResult.Written.class, name = "WRITTEN"),
    @JsonSubTypes.Type(value = WriteResult.NotWritten.class, name = "NOT_WRITTEN")
  })
  sealed interface WriteResult permits WriteResult.Written, WriteResult.NotWritten {
    /** Workbook file was written successfully. */
    record Written(String executionPath) implements WriteResult {
      public Written {
        Objects.requireNonNull(executionPath, "executionPath must not be null");
        if (executionPath.isBlank()) {
          throw new IllegalArgumentException("executionPath must not be blank");
        }
      }
    }

    /** Workbook file was not written before the run failed. */
    record NotWritten() implements WriteResult {}
  }
}
