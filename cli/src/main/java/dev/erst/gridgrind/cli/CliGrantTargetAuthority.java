package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.util.List;

/** CLI-owned workbook-target authority translated into the engine grant boundary. */
@Union(discriminator = "type")
public sealed interface CliGrantTargetAuthority
    permits CliGrantTargetAuthority.WorkbookWide, CliGrantTargetAuthority.SelectedSheets {
  /** Translates this CLI target authority into the engine-owned value. */
  GridGrindExecutionGrant.WorkbookTargetAuthority toEngineAuthority();

  /** Grants workbook-wide operation footprints. */
  @Variant(id = "WORKBOOK_WIDE", summary = "Grants workbook-wide operation footprints.")
  record WorkbookWide() implements CliGrantTargetAuthority {
    @Override
    public GridGrindExecutionGrant.WorkbookTargetAuthority toEngineAuthority() {
      return new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide();
    }
  }

  /** Grants only listed sheets when an operation footprint is provably bounded. */
  @Variant(
      id = "SELECTED_SHEETS",
      summary = "Grants only listed sheets when an operation footprint is provably bounded.")
  record SelectedSheets(List<String> sheetNames) implements CliGrantTargetAuthority {
    public SelectedSheets {
      sheetNames = CliGrantValueSupport.copyStrings(sheetNames, "sheetNames");
    }

    @Override
    public GridGrindExecutionGrant.WorkbookTargetAuthority toEngineAuthority() {
      return new GridGrindExecutionGrant.WorkbookTargetAuthority.SelectedSheets(sheetNames);
    }
  }
}
