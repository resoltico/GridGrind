package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;

/** Workbook-persistence descriptors for the public protocol catalog. */
final class GridGrindProtocolCatalogPersistenceTypeDescriptors {
  static final List<CatalogTypeDescriptor> PERSISTENCE_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookPersistence.None.class,
              "NONE",
              "Keep the workbook in memory only." + " The response persistence.type echoes NONE."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookPersistence.OverwriteSource.class,
              "OVERWRITE",
              "Overwrite the opened source workbook at source.path."
                  + " No path field is accepted on OVERWRITE;"
                  + " the write target is the same path opened by the EXISTING source."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary()
                  + " persistence.security can encrypt and/or sign the saved OOXML package."
                  + " The response persistence.type echoes OVERWRITE and includes sourcePath"
                  + " (the original source path string) and executionPath (absolute normalized).",
              "security"),
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookPersistence.SaveAs.class,
              "SAVE_AS",
              "Save the workbook to a new .xlsx path."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary()
                  + " persistence.security can encrypt and/or sign the saved OOXML package."
                  + " The response persistence.type echoes SAVE_AS and includes requestedPath"
                  + " (the literal path from the request) and executionPath (the absolute"
                  + " normalized path where the file was written); they differ when a relative"
                  + " path or a path with .. segments is supplied."
                  + " Missing parent directories are created automatically.",
              "security"));

  private GridGrindProtocolCatalogPersistenceTypeDescriptors() {}
}
