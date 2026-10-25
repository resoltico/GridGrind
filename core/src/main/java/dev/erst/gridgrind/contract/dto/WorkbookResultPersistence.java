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

    /** Workbook remained in memory only and was not published to a destination. */
    record NotSaved() implements PersistenceOutcome {}

    /** Workbook targeted the path supplied in the SAVE_AS persistence field. */
    record SavedAs(String requestedPath, PublicationOutcome publication)
        implements PersistenceOutcome {
      public SavedAs {
        requestedPath = requireNonBlank(requestedPath, "requestedPath");
        Objects.requireNonNull(publication, "publication must not be null");
      }
    }

    /** Workbook targeted the opened source workbook path for overwrite persistence. */
    record Overwritten(
        @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> sourcePath,
        PublicationOutcome publication)
        implements PersistenceOutcome {
      public Overwritten {
        sourcePath = normalizeSourcePath(sourcePath);
        Objects.requireNonNull(publication, "publication must not be null");
      }

      /** Creates one overwrite outcome that echoes a known EXISTING source path. */
      public Overwritten(String sourcePath, PublicationOutcome publication) {
        this(Optional.of(sourcePath), publication);
      }

      private static Optional<String> normalizeSourcePath(Optional<String> sourcePath) {
        Optional<String> normalized = Objects.requireNonNullElseGet(sourcePath, Optional::empty);
        if (normalized.isEmpty()) {
          return Optional.empty();
        }
        return Optional.of(requireNonBlank(normalized.orElseThrow(), "sourcePath"));
      }
    }
  }

  /** Reports what GridGrind established about one requested final workbook destination. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = PublicationOutcome.NotAttempted.class, name = "NOT_ATTEMPTED"),
    @JsonSubTypes.Type(value = PublicationOutcome.NotPublished.class, name = "NOT_PUBLISHED"),
    @JsonSubTypes.Type(value = PublicationOutcome.Published.class, name = "PUBLISHED"),
    @JsonSubTypes.Type(value = PublicationOutcome.Uncertain.class, name = "UNCERTAIN")
  })
  sealed interface PublicationOutcome
      permits PublicationOutcome.NotAttempted,
          PublicationOutcome.NotPublished,
          PublicationOutcome.Published,
          PublicationOutcome.Uncertain {

    /** Final destination mutation was never attempted. */
    record NotAttempted() implements PublicationOutcome {}

    /**
     * Final destination remains provably absent or unchanged after publication did not complete.
     */
    record NotPublished(DestinationState destinationState) implements PublicationOutcome {
      public NotPublished {
        Objects.requireNonNull(destinationState, "destinationState must not be null");
      }
    }

    /** Final destination was observed with the complete verified staged artifact. */
    record Published(
        String executionPath,
        String sha256,
        long byteSize,
        StagedArtifactVerification stagedArtifactVerification,
        DurabilityEvidence durability)
        implements PublicationOutcome {
      public Published {
        executionPath = requireNonBlank(executionPath, "executionPath");
        sha256 = requireSha256(sha256);
        if (byteSize < 0) {
          throw new IllegalArgumentException("byteSize must be >= 0");
        }
        Objects.requireNonNull(
            stagedArtifactVerification, "stagedArtifactVerification must not be null");
        Objects.requireNonNull(durability, "durability must not be null");
      }
    }

    /** Final destination may have changed, so automated retry is unsafe without inspection. */
    record Uncertain(String executionPath, Recovery recovery) implements PublicationOutcome {
      public Uncertain {
        executionPath = requireNonBlank(executionPath, "executionPath");
        Objects.requireNonNull(recovery, "recovery must not be null");
      }
    }

    /** Known final-destination state after a publication failure. */
    enum DestinationState {
      ABSENT,
      PRESERVED
    }

    /** Required operator response to an uncertain publication attempt. */
    enum Recovery {
      INSPECT_DESTINATION_DO_NOT_RETRY
    }

    /** Staged artifact successfully reopened with its effective package-security settings. */
    record StagedArtifactVerification() {}

    /** Durability evidence established for one published artifact. */
    @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "level")
    @JsonSubTypes({
      @JsonSubTypes.Type(
          value = DurabilityEvidence.FileAndDirectorySynced.class,
          name = "FILE_AND_DIRECTORY_SYNCED"),
      @JsonSubTypes.Type(
          value = DurabilityEvidence.FileSyncedDirectoryUnestablished.class,
          name = "FILE_SYNCED_DIRECTORY_UNESTABLISHED")
    })
    sealed interface DurabilityEvidence
        permits DurabilityEvidence.FileAndDirectorySynced,
            DurabilityEvidence.FileSyncedDirectoryUnestablished {
      /** File data and the containing directory entry were synchronized. */
      record FileAndDirectorySynced() implements DurabilityEvidence {}

      /** File data was synchronized but directory synchronization was not established. */
      record FileSyncedDirectoryUnestablished() implements DurabilityEvidence {}
    }

    private static String requireSha256(String value) {
      String normalized = requireNonBlank(value, "sha256");
      if (!normalized.matches("[0-9a-f]{64}")) {
        throw new IllegalArgumentException("sha256 must be 64 lowercase hexadecimal characters");
      }
      return normalized;
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
