package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;

/** Workbook-source descriptors for the public protocol catalog. */
final class GridGrindProtocolCatalogSourceTypeDescriptors {
  static final List<CatalogTypeDescriptor> SOURCE_TYPES =
      List.of(
          CatalogTypeEntryFactory.descriptor(
              WorkbookPlan.WorkbookSource.New.class,
              "NEW",
              "Create a brand-new empty workbook. A new workbook starts with zero sheets;"
                  + " use ENSURE_SHEET to create the first sheet."),
          CatalogTypeEntryFactory.descriptorWithNotes(
              WorkbookPlan.WorkbookSource.ExistingFile.class,
              "EXISTING",
              "Open an existing .xlsx workbook from disk."
                  + " source.security.password unlocks encrypted OOXML packages.",
              GridGrindProtocolCatalogNotes.requestOwnedPathRuleRef(),
              "security"));

  private GridGrindProtocolCatalogSourceTypeDescriptors() {}
}
