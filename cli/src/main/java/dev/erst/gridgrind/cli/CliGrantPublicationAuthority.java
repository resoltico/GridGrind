package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.nio.file.Path;
import java.util.Objects;

/** CLI-owned final-artifact authority translated into the engine grant boundary. */
@Union(discriminator = "type")
public sealed interface CliGrantPublicationAuthority
    permits CliGrantPublicationAuthority.None,
        CliGrantPublicationAuthority.SaveAs,
        CliGrantPublicationAuthority.OverwriteSource {
  /** Converts this document-owned authority using the CLI process working directory as its root. */
  GridGrindExecutionGrant.PublicationAuthority toEngineAuthority(Path cliWorkingDirectory);

  @Variant(id = "NONE", summary = "Grant no workbook publication authority.")
  record None() implements CliGrantPublicationAuthority {
    @Override
    public GridGrindExecutionGrant.PublicationAuthority toEngineAuthority(
        Path cliWorkingDirectory) {
      Objects.requireNonNull(cliWorkingDirectory, "cliWorkingDirectory must not be null");
      return new GridGrindExecutionGrant.PublicationAuthority.None();
    }
  }

  @Variant(id = "SAVE_AS", summary = "Grant publication to one explicit output path.")
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

  @Variant(id = "OVERWRITE_SOURCE", summary = "Grant replacement of the existing workbook source.")
  record OverwriteSource() implements CliGrantPublicationAuthority {
    @Override
    public GridGrindExecutionGrant.PublicationAuthority toEngineAuthority(
        Path cliWorkingDirectory) {
      Objects.requireNonNull(cliWorkingDirectory, "cliWorkingDirectory must not be null");
      return new GridGrindExecutionGrant.PublicationAuthority.OverwriteSource();
    }
  }
}
