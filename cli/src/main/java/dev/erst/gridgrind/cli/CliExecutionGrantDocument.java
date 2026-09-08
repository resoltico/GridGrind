package dev.erst.gridgrind.cli;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.SecretReference;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** CLI-owned host grant document translated into the transport-neutral engine authority model. */
public record CliExecutionGrantDocument(
    List<ReadableResource> readableResources,
    List<String> operationIds,
    CliGrantTargetAuthority targetAuthority,
    CliGrantPublicationAuthority publicationAuthority,
    List<SecretReference> allowedSecretReferences,
    CliGrantAcceptancePolicy acceptancePolicy) {
  public CliExecutionGrantDocument {
    readableResources = CliGrantValueSupport.copyValues(readableResources, "readableResources");
    operationIds = CliGrantValueSupport.copyStrings(operationIds, "operationIds");
    Objects.requireNonNull(targetAuthority, "targetAuthority must not be null");
    Objects.requireNonNull(publicationAuthority, "publicationAuthority must not be null");
    allowedSecretReferences =
        CliGrantValueSupport.copyValues(allowedSecretReferences, "allowedSecretReferences");
    Objects.requireNonNull(acceptancePolicy, "acceptancePolicy must not be null");
  }

  /** Converts this CLI document using the process working directory as its path base. */
  GridGrindExecutionGrant toExecutionGrant(Path cliWorkingDirectory) {
    Path base =
        Objects.requireNonNull(cliWorkingDirectory, "cliWorkingDirectory must not be null")
            .toAbsolutePath()
            .normalize();
    return new GridGrindExecutionGrant.Bounded(
        readableResources.stream().map(resource -> resource.toEngineAuthority(base)).toList(),
        operationIds,
        targetAuthority.toEngineAuthority(),
        publicationAuthority.toEngineAuthority(base),
        allowedSecretReferences,
        acceptancePolicy.toEnginePolicy());
  }

  /** One resource the CLI host explicitly permits the plan to read. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = ReadableResource.File.class, name = "FILE"),
    @JsonSubTypes.Type(value = ReadableResource.StandardInput.class, name = "STANDARD_INPUT")
  })
  public sealed interface ReadableResource
      permits ReadableResource.File, ReadableResource.StandardInput {
    /** Translates one CLI resource into the engine-owned authority value. */
    GridGrindExecutionGrant.ReadAuthority toEngineAuthority(Path cliWorkingDirectory);

    /**
     * Grants one exact file resource, rooted at the CLI process working directory when relative.
     */
    record File(String path) implements ReadableResource {
      public File {
        path = CliGrantValueSupport.requireNonBlank(path, "path");
      }

      @Override
      public GridGrindExecutionGrant.ReadAuthority toEngineAuthority(Path cliWorkingDirectory) {
        return new GridGrindExecutionGrant.ReadAuthority.File(
            CliGrantValueSupport.resolve(path, cliWorkingDirectory));
      }
    }

    /** Grants the standard-input payload already bound by the CLI host. */
    record StandardInput() implements ReadableResource {
      @Override
      public GridGrindExecutionGrant.ReadAuthority toEngineAuthority(Path cliWorkingDirectory) {
        Objects.requireNonNull(cliWorkingDirectory, "cliWorkingDirectory must not be null");
        return new GridGrindExecutionGrant.ReadAuthority.StandardInput();
      }
    }
  }
}
