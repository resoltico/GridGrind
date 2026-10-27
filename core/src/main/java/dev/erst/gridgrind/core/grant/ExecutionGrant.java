package dev.erst.gridgrind.core.grant;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.SecretReference;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Explicit authority for one execution; workbook plans cannot create this grant. */
public record ExecutionGrant(
    List<ReadAuthority> readableResources,
    List<String> operationIds,
    WorkbookTargetAuthority targetAuthority,
    PublicationAuthority publicationAuthority,
    List<SecretReference> allowedSecretReferences,
    AcceptancePolicy acceptancePolicy) {
  public ExecutionGrant {
    readableResources = copyReadableResources(readableResources);
    operationIds = copyOperationIds(operationIds);
    Objects.requireNonNull(targetAuthority, "targetAuthority must not be null");
    Objects.requireNonNull(publicationAuthority, "publicationAuthority must not be null");
    allowedSecretReferences = copySecretReferences(allowedSecretReferences);
    Objects.requireNonNull(acceptancePolicy, "acceptancePolicy must not be null");
  }

  private static List<ReadAuthority> copyReadableResources(List<ReadAuthority> readableResources) {
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
        throw new IllegalArgumentException("operationIds must contain SCREAMING_SNAKE_CASE values");
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

  /** Returns explicit authority permitting no operations, reads, secrets or publication. */
  public static ExecutionGrant denied() {
    return new ExecutionGrant(
        List.of(),
        List.of(),
        new WorkbookTargetAuthority.WorkbookWide(),
        new PublicationAuthority.None(),
        List.of(),
        AcceptancePolicy.minimum());
  }

  /** One resource the trusted host permits the execution to read. */
  @Union(discriminator = "type")
  public sealed interface ReadAuthority permits ReadAuthority.File, ReadAuthority.Stdin {
    /** Grants reading exactly one normalized file resource. */
    @Variant(id = "FILE", summary = "Grant reading one exact normalized file resource.")
    record File(Path path) implements ReadAuthority {
      public File {
        path = Objects.requireNonNull(path, "path must not be null").toAbsolutePath().normalize();
      }
    }

    /** Grants consumption of the host-bound standard-input payload. */
    @Variant(
        id = "STANDARD_INPUT",
        summary = "Grant consumption of the bound standard-input payload.")
    record Stdin() implements ReadAuthority {}
  }

  /** Workbook scope the host permits a resolved operation footprint to affect or inspect. */
  @Union(discriminator = "type")
  public sealed interface WorkbookTargetAuthority
      permits WorkbookTargetAuthority.WorkbookWide, WorkbookTargetAuthority.SelectedSheets {
    /** Grants workbook-wide operation effects. */
    @Variant(id = "WORKBOOK_WIDE", summary = "Grant workbook-wide operation footprints.")
    record WorkbookWide() implements WorkbookTargetAuthority {}

    /** Grants only explicitly named sheet targets when their footprint is provably bounded. */
    @Variant(
        id = "SELECTED_SHEETS",
        summary = "Grant only listed sheets when an operation footprint is provably bounded.")
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
  @Union(discriminator = "type")
  public sealed interface PublicationAuthority
      permits PublicationAuthority.None,
          PublicationAuthority.SaveAs,
          PublicationAuthority.OverwriteSource {
    /** Grants no final workbook publication. */
    @Variant(id = "NONE", summary = "Grant no workbook publication authority.")
    record None() implements PublicationAuthority {}

    /** Grants one exact SAVE_AS target and collision policy. */
    @Variant(
        id = "SAVE_AS",
        summary = "Grant publication to one exact output path and collision policy.")
    record SaveAs(Path path, WorkbookPlan.WorkbookPersistence.IfExists ifExists)
        implements PublicationAuthority {
      public SaveAs {
        path = Objects.requireNonNull(path, "path must not be null").toAbsolutePath().normalize();
        Objects.requireNonNull(ifExists, "ifExists must not be null");
      }
    }

    /** Grants overwriting exactly the admitted existing workbook source. */
    @Variant(
        id = "OVERWRITE_SOURCE",
        summary = "Grant replacement of the admitted existing workbook source.")
    record OverwriteSource() implements PublicationAuthority {}
  }
}
