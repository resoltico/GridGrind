package dev.erst.gridgrind.cli;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.util.List;

/** CLI-owned workbook-target authority translated into the engine grant boundary. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CliGrantTargetAuthority.WorkbookWide.class, name = "WORKBOOK_WIDE"),
  @JsonSubTypes.Type(value = CliGrantTargetAuthority.SelectedSheets.class, name = "SELECTED_SHEETS")
})
public sealed interface CliGrantTargetAuthority
    permits CliGrantTargetAuthority.WorkbookWide, CliGrantTargetAuthority.SelectedSheets {
  /** Translates this CLI target authority into the engine-owned value. */
  GridGrindExecutionGrant.WorkbookTargetAuthority toEngineAuthority();

  /** Grants workbook-wide operation footprints. */
  record WorkbookWide() implements CliGrantTargetAuthority {
    @Override
    public GridGrindExecutionGrant.WorkbookTargetAuthority toEngineAuthority() {
      return new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide();
    }
  }

  /** Grants only listed sheets when an operation footprint is provably bounded. */
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
