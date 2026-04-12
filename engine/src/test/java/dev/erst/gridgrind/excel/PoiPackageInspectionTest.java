package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.junit.jupiter.api.Test;

/** Tests the package-inspection seam used by pivot package cleanup and allocation logic. */
class PoiPackageInspectionTest {
  @Test
  void packageInspectionCopiesNormalResultsAndWrapsInvalidFormatFailures() throws Exception {
    try (OPCPackage pkg = OPCPackage.create(new java.io.ByteArrayOutputStream())) {
      PackagePart part =
          pkg.createPart(
              org.apache.poi.openxml4j.opc.PackagingURIHelper.createPartName("/xl/test.xml"),
              "application/xml");

      assertEquals(List.of(part), PoiPackageInspection.packageParts(() -> List.of(part), "parts"));
      assertEquals(List.of(), PoiPackageInspection.relationships(() -> List.of(), "relationships"));
      assertNotNull(PoiPackageInspection.packageParts(pkg, "parts"));
      assertEquals(List.of(), PoiPackageInspection.relationships(part, "relationships"));
    }

    IllegalStateException partFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                PoiPackageInspection.packageParts(
                    () -> {
                      throw new InvalidFormatException("boom");
                    },
                    "parts"));
    assertEquals("parts", partFailure.getMessage());
    assertInstanceOf(InvalidFormatException.class, partFailure.getCause());

    IllegalStateException relationshipFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                PoiPackageInspection.relationships(
                    () -> {
                      throw new InvalidFormatException("boom");
                    },
                    "relationships"));
    assertEquals("relationships", relationshipFailure.getMessage());
    assertInstanceOf(InvalidFormatException.class, relationshipFailure.getCause());
  }
}
