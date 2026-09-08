package dev.erst.gridgrind.cli;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.nio.file.Path;
import java.util.Objects;

/** CLI-owned final-artifact authority translated into the engine grant boundary. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CliGrantPublicationAuthority.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = CliGrantPublicationAuthority.SaveAs.class, name = "SAVE_AS"),
  @JsonSubTypes.Type(
      value = CliGrantPublicationAuthority.OverwriteSource.class,
      name = "OVERWRITE_SOURCE")
})
public sealed interface CliGrantPublicationAuthority
    permits CliGrantPublicationAuthority.None,
        CliGrantPublicationAuthority.SaveAs,
        CliGrantPublicationAuthority.OverwriteSource {
  /** Converts this document-owned authority using the CLI process working directory as its root. */
  GridGrindExecutionGrant.PublicationAuthority toEngineAuthority(Path cliWorkingDirectory);

  record None() implements CliGrantPublicationAuthority {
    @Override
    public GridGrindExecutionGrant.PublicationAuthority toEngineAuthority(
        Path cliWorkingDirectory) {
      Objects.requireNonNull(cliWorkingDirectory, "cliWorkingDirectory must not be null");
      return new GridGrindExecutionGrant.PublicationAuthority.None();
    }
  }

  record SaveAs(String path, WorkbookPlan.WorkbookPersistence.IfExists ifExists)
      implements CliGrantPublicationAuthority {
    public SaveAs {
      path = CliGrantValueSupport.requireNonBlank(path, "path");
      Objects.requireNonNull(ifExists, "ifExists must not be null");
    }

    @Override
    public GridGrindExecutionGrant.PublicationAuthority toEngineAuthority(
        Path cliWorkingDirectory) {
      return new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
          CliGrantValueSupport.resolve(path, cliWorkingDirectory), ifExists);
    }
  }

  record OverwriteSource() implements CliGrantPublicationAuthority {
    @Override
    public GridGrindExecutionGrant.PublicationAuthority toEngineAuthority(
        Path cliWorkingDirectory) {
      Objects.requireNonNull(cliWorkingDirectory, "cliWorkingDirectory must not be null");
      return new GridGrindExecutionGrant.PublicationAuthority.OverwriteSource();
    }
  }
}
