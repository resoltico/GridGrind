package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTHyperlink;

/** Materializes POI sheet-clone prerequisites that are otherwise deferred until save-time. */
final class ExcelSheetClonePreparationSupport {
  /** Ensures a source sheet is internally consistent before handing it to POI clone logic. */
  void prepareSourceSheetForClone(XSSFSheet sourceSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    materializeExternalHyperlinkRelationships(sourceSheet);
  }

  private static void materializeExternalHyperlinkRelationships(XSSFSheet sourceSheet) {
    for (XSSFHyperlink hyperlink : List.copyOf(sourceSheet.getHyperlinkList())) {
      materializeExternalHyperlinkRelationship(sourceSheet, hyperlink)
          .ifPresent(
              materialized -> {
                sourceSheet.removeHyperlink(hyperlink);
                sourceSheet.addHyperlink(materialized);
              });
    }
  }

  private static Optional<XSSFHyperlink> materializeExternalHyperlinkRelationship(
      XSSFSheet sourceSheet, XSSFHyperlink hyperlink) {
    if (!hyperlink.needsRelationToo()) {
      return Optional.empty();
    }
    String relationshipId = hyperlink.getCTHyperlink().getId();
    if (relationshipId != null
        && !relationshipId.isBlank()
        && sourceSheet.getPackagePart().getRelationship(relationshipId) != null) {
      return Optional.empty();
    }
    return Optional.of(
        MaterializedHyperlink.wrapWithMaterializedRelation(
            hyperlink, sourceSheet.getPackagePart()));
  }

  /**
   * POI cloneSheet consumes relation-backed XSSFHyperlink objects rather than raw CT hyperlink ids.
   * Rehydrating through the public copy constructor and protected relation materializer keeps the
   * hyperlink list and package relationships aligned without private access.
   */
  private static final class MaterializedHyperlink extends XSSFHyperlink {
    private MaterializedHyperlink(CTHyperlink ctHyperlink, PackageRelationship relationship) {
      super(ctHyperlink, relationship);
    }

    private static XSSFHyperlink wrapWithMaterializedRelation(
        XSSFHyperlink hyperlink, PackagePart sheetPart) {
      String relationshipId = hyperlink.getCTHyperlink().getId();
      String requestedRelationshipId =
          relationshipId == null || relationshipId.isBlank() ? null : relationshipId;
      PackageRelationship relationship =
          sheetPart.addExternalRelationship(
              hyperlink.getAddress(),
              XSSFRelation.SHEET_HYPERLINKS.getRelation(),
              requestedRelationshipId);
      hyperlink.getCTHyperlink().setId(relationship.getId());
      return new MaterializedHyperlink(hyperlink.getCTHyperlink(), relationship);
    }
  }
}
