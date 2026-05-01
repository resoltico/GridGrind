package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;

/** Workbook-source descriptors for the public protocol catalog. */
final class GridGrindProtocolCatalogSourceTypeDescriptors {
  static final List<CatalogTypeDescriptor> SOURCE_TYPES =
      List.of(
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookSource.New.class,
              "NEW",
              "Create a brand-new empty workbook. A new workbook starts with zero sheets;"
                  + " use ENSURE_SHEET to create the first sheet."),
          GridGrindProtocolCatalog.descriptor(
              WorkbookPlan.WorkbookSource.ExistingFile.class,
              "EXISTING",
              "Open an existing .xlsx workbook from disk."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary()
                  + " source.security.password unlocks encrypted OOXML packages.",
              "security"));

  private GridGrindProtocolCatalogSourceTypeDescriptors() {}
}
