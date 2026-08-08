package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;

/** Workbook-persistence descriptors for the public protocol catalog. */
final class GridGrindProtocolCatalogPersistenceTypeDescriptors {
  static final List<CatalogTypeDescriptor> PERSISTENCE_TYPES =
      List.of(
          CatalogTypeEntryFactory.descriptor(
              WorkbookPlan.WorkbookPersistence.None.class,
              "NONE",
              "Keep the workbook in memory only." + " The response persistence.type echoes NONE."),
          CatalogTypeEntryFactory.descriptorWithNotes(
              WorkbookPlan.WorkbookPersistence.Overwrite.class,
              "OVERWRITE",
              "Overwrite the opened source workbook at source.path."
                  + " No path field is accepted on OVERWRITE;"
                  + " the write target is the same path opened by the EXISTING source."
                  + " persistence.security can encrypt and/or sign the saved OOXML package."
                  + " The response persistence.type echoes OVERWRITE, includes sourcePath"
                  + " (the original source path string) whenever an EXISTING source path was"
                  + " available, and otherwise omits sourcePath instead of inventing one."
                  + " It carries write.status=WRITTEN with executionPath after a successful save"
                  + " or write.status=NOT_WRITTEN when the run fails before any file is updated.",
              GridGrindProtocolCatalogNotes.requestOwnedPathRuleRef()),
          CatalogTypeEntryFactory.descriptorWithNotes(
              WorkbookPlan.WorkbookPersistence.SaveAs.class,
              "SAVE_AS",
              "Save the workbook to one .xlsx path with an explicit ifExists collision policy."
                  + " ifExists=REJECT requires the target path to be absent."
                  + " ifExists=REPLACE enables create-or-replace."
                  + " persistence.security can encrypt and/or sign the saved OOXML package."
                  + " The response persistence.type echoes SAVE_AS, includes requestedPath"
                  + " (the literal path from the request), and carries write.status=WRITTEN with"
                  + " executionPath (the absolute normalized path where the file was written) or"
                  + " write.status=NOT_WRITTEN when the run fails before the target file is"
                  + " created."
                  + " Missing parent directories are created automatically.",
              GridGrindProtocolCatalogNotes.requestOwnedPathRuleRef()));

  private GridGrindProtocolCatalogPersistenceTypeDescriptors() {}
}
