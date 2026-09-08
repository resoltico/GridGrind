package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.SecretReference;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Trusted-host authority for one GridGrind execution; workbook plans cannot create this grant. */
public sealed interface GridGrindExecutionGrant permits GridGrindExecutionGrant.Bounded {
  /** Explicit authority bounded to declared reads, operations, workbook scope, and publication. */
  record Bounded(
      List<ReadAuthority> readableResources,
      List<String> operationIds,
      WorkbookTargetAuthority workbookTargetAuthority,
      PublicationAuthority publicationAuthority,
      List<SecretReference> allowedSecretReferences,
      GridGrindHostAcceptancePolicy hostAcceptancePolicy)
      implements GridGrindExecutionGrant {
    public Bounded {
      readableResources = copyReadableResources(readableResources);
      operationIds = copyOperationIds(operationIds);
      Objects.requireNonNull(workbookTargetAuthority, "workbookTargetAuthority must not be null");
      Objects.requireNonNull(publicationAuthority, "publicationAuthority must not be null");
      allowedSecretReferences = copySecretReferences(allowedSecretReferences);
      Objects.requireNonNull(hostAcceptancePolicy, "hostAcceptancePolicy must not be null");
    }

    private static List<ReadAuthority> copyReadableResources(
        List<ReadAuthority> readableResources) {
      Objects.requireNonNull(readableResources, "readableResources must not be null");
      Set<ReadAuthority> unique = new LinkedHashSet<>();
      for (ReadAuthority resource : readableResources) {
        unique.add(
            Objects.requireNonNull(resource, "readableResources must not contain null values"));
      }
      return List.copyOf(unique);
    }

    private static List<String> copyOperationIds(List<String> operationIds) {
      Objects.requireNonNull(operationIds, "operationIds must not be null");
      Set<String> unique = new LinkedHashSet<>();
      for (String operationId : operationIds) {
        String normalized =
            Objects.requireNonNull(operationId, "operationIds must not contain null values");
        if (!normalized.matches("[A-Z][A-Z0-9_]*")) {
          throw new IllegalArgumentException(
              "operationIds must contain SCREAMING_SNAKE_CASE values");
        }
        unique.add(normalized);
      }
      return List.copyOf(unique);
    }

    private static List<SecretReference> copySecretReferences(
        List<SecretReference> allowedSecretReferences) {
      Objects.requireNonNull(allowedSecretReferences, "allowedSecretReferences must not be null");
      return List.copyOf(new LinkedHashSet<>(allowedSecretReferences));
    }
  }

  /** One resource the trusted host permits the execution to read. */
  sealed interface ReadAuthority permits ReadAuthority.File, ReadAuthority.StandardInput {
    /** Grants reading exactly one normalized file resource. */
    record File(Path path) implements ReadAuthority {
      public File {
        path = Objects.requireNonNull(path, "path must not be null").toAbsolutePath().normalize();
      }
    }

    /** Grants consumption of the host-bound standard-input payload. */
    record StandardInput() implements ReadAuthority {}
  }

  /** Workbook scope the host permits a resolved operation footprint to affect or inspect. */
  sealed interface WorkbookTargetAuthority
      permits WorkbookTargetAuthority.WorkbookWide, WorkbookTargetAuthority.SelectedSheets {
    /** Grants workbook-wide operation effects. */
    record WorkbookWide() implements WorkbookTargetAuthority {}

    /** Grants only explicitly named sheet targets when their footprint is provably bounded. */
    record SelectedSheets(List<String> sheetNames) implements WorkbookTargetAuthority {
      public SelectedSheets {
        Objects.requireNonNull(sheetNames, "sheetNames must not be null");
        List<String> copied = new ArrayList<>(sheetNames.size());
        for (String sheetName : sheetNames) {
          String normalized =
              Objects.requireNonNull(sheetName, "sheetNames must not contain null values");
          if (normalized.isBlank()) {
            throw new IllegalArgumentException("sheetNames must not contain blank values");
          }
          copied.add(normalized);
        }
        sheetNames = List.copyOf(new LinkedHashSet<>(copied));
      }
    }
  }

  /** Publication authority for the final workbook artifact. */
  sealed interface PublicationAuthority
      permits PublicationAuthority.None,
          PublicationAuthority.SaveAs,
          PublicationAuthority.OverwriteSource {
    /** Grants no final workbook publication. */
    record None() implements PublicationAuthority {}

    /** Grants one exact SAVE_AS target and collision policy. */
    record SaveAs(Path path, WorkbookPlan.WorkbookPersistence.IfExists ifExists)
        implements PublicationAuthority {
      public SaveAs {
        path = Objects.requireNonNull(path, "path must not be null").toAbsolutePath().normalize();
        Objects.requireNonNull(ifExists, "ifExists must not be null");
      }
    }

    /** Grants overwriting exactly the admitted existing workbook source. */
    record OverwriteSource() implements PublicationAuthority {}
  }
}
