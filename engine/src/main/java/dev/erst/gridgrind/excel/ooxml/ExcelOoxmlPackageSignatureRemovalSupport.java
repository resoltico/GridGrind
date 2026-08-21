package dev.erst.gridgrind.excel.ooxml;

import dev.erst.gridgrind.excel.WorkbookSecurityException;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;

/** Removes OOXML digital-signature parts before an unsigned save or fresh package signature. */
final class ExcelOoxmlPackageSignatureRemovalSupport {
  private static final String DIGITAL_SIGNATURE_ORIGIN_RELATIONSHIP =
      "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
  private static final String SIGNATURE_PART_DIRECTORY = "/_xmlsignatures/";

  private ExcelOoxmlPackageSignatureRemovalSupport() {}

  static void removeSignatures(Path workbookPath) throws IOException {
    try (OPCPackage pkg = OPCPackage.open(workbookPath.toFile(), PackageAccess.READ_WRITE)) {
      List<String> originRelationshipIds = new ArrayList<>();
      for (PackageRelationship relationship : pkg.getRelationships()) {
        if (DIGITAL_SIGNATURE_ORIGIN_RELATIONSHIP.equals(relationship.getRelationshipType())) {
          originRelationshipIds.add(relationship.getId());
        }
      }
      List<PackagePartName> signaturePartNames =
          pkg.getParts().stream()
              .map(part -> part.getPartName())
              .filter(partName -> partName.getName().startsWith(SIGNATURE_PART_DIRECTORY))
              .toList();
      for (PackagePartName signaturePartName : signaturePartNames) {
        pkg.removePart(signaturePartName);
      }
      for (String originRelationshipId : originRelationshipIds) {
        pkg.removeRelationship(originRelationshipId);
      }
    } catch (InvalidFormatException | NotOfficeXmlFileException exception) {
      throw new WorkbookSecurityException(
          "Failed to open the OOXML workbook package for signature removal: " + workbookPath,
          exception);
    }
  }
}
