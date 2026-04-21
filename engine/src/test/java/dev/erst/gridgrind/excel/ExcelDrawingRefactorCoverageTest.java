package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused coverage for drawing helpers extracted during controller/chart breakup. */
class ExcelDrawingRefactorCoverageTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  private static final byte[] INVALID_PNG_BYTES =
      Base64.getDecoder().decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVQ=");

  @Test
  void extractedControllerFacadesDelegateWithoutReinflatingTheController() throws Throwable {
    ExcelDrawingController controller = new ExcelDrawingController();

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      XSSFObjectData objectData =
          createEmbeddedObject(workbook, sheet.createDrawingPatriarch(), "OpsEmbed", 0, 0, 3, 3);

      assertEquals("OpsEmbed", controller.snapshotEmbeddedObject(objectData).name());
      assertNotNull(controller.oleObjectPart(objectData));
      assertNull(controller.nullIfBlank(" "));

      ExcelDrawingController.RasterDimensions dimensions =
          ExcelDrawingController.rasterDimensions(PNG_PIXEL_BYTES);
      assertEquals(1, dimensions.widthPixels());
      assertEquals(1, dimensions.heightPixels());
    }
  }

  @Test
  void extractedChartAndSnapshotHelpersKeepFacadeAndFailureBranchesCovered() throws Throwable {
    assertEquals(
        "'Data Sheet'!$A$1:$B$2",
        ExcelDrawingChartSupport.normalizeAreaFormulaForPoi("'Data Sheet'!$A$1:'Data Sheet'!$B$2"));

    ExcelDrawingSnapshotSupport.RasterDimensions noImageDimensions =
        ExcelDrawingSnapshotSupport.rasterDimensions(
            "not-an-image".getBytes(StandardCharsets.UTF_8));
    assertNull(noImageDimensions.widthPixels());
    assertNull(noImageDimensions.heightPixels());

    ExcelDrawingSnapshotSupport.RasterDimensions invalidPngDimensions =
        ExcelDrawingSnapshotSupport.rasterDimensions(INVALID_PNG_BYTES);
    assertNull(invalidPngDimensions.widthPixels());
    assertNull(invalidPngDimensions.heightPixels());
  }

  @Test
  void binaryCleanupSupportCoversNullUsedMissingDeletedAndInvalidPackages() throws Exception {
    assertFalse(ExcelPackageRelationshipSupport.partIsStillReferenced(List.of(), null));

    try (OPCPackage pkg = OPCPackage.create(new ByteArrayOutputStream())) {
      PackagePart source =
          pkg.createPart(
              PackagingURIHelper.createPartName("/xl/workbook.xml"),
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
      PackagePart usedPart =
          pkg.createPart(PackagingURIHelper.createPartName("/xl/media/used.png"), "image/png");
      PackagePart unusedPart =
          pkg.createPart(PackagingURIHelper.createPartName("/xl/media/unused.png"), "image/png");

      pkg.createPart(
          PackagingURIHelper.createPartName("/xl/_rels/workbook.xml.rels"),
          "application/vnd.openxmlformats-package.relationships+xml");

      source.addRelationship(
          usedPart.getPartName(), TargetMode.INTERNAL, "urn:gridgrind:test", "rIdUsed");
      source.addRelationship(
          unusedPart.getPartName(), TargetMode.INTERNAL, "urn:gridgrind:test", "rIdUnused");
      source.addExternalRelationship(
          "https://example.com/object", "urn:gridgrind:test", "rIdExternal");

      ExcelDrawingBinarySupport.cleanupPackagePartIfUnused(pkg, null);
      assertTrue(pkg.containPart(usedPart.getPartName()));

      ExcelDrawingBinarySupport.cleanupPackagePartIfUnused(pkg, usedPart.getPartName());
      assertTrue(pkg.containPart(usedPart.getPartName()));

      ExcelDrawingBinarySupport.cleanupPackagePartIfUnused(
          pkg, PackagingURIHelper.createPartName("/xl/media/missing.png"));

      ExcelDrawingBinarySupport.removeRelationshipsToPart(source, unusedPart.getPartName());
      ExcelDrawingBinarySupport.cleanupPackagePartIfUnused(pkg, unusedPart.getPartName());
      assertFalse(pkg.containPart(unusedPart.getPartName()));
    }

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelPackageRelationshipSupport.partIsStillReferenced(
                    List.of(
                        new InvalidRelationshipsPackagePart(
                            "/xl/worksheets/sheet1.xml",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")),
                    PackagingURIHelper.createPartName("/xl/media/unused.png")));
    assertTrue(failure.getMessage().contains("Failed to inspect package relationships"));

    IllegalStateException invalidPackageFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelPackageRelationshipSupport.requireParts(
                    () -> {
                      throw new InvalidFormatException("broken package");
                    }));
    assertTrue(
        invalidPackageFailure.getMessage().contains("Failed to inspect package relationships"));
  }

  private static XSSFObjectData createEmbeddedObject(
      XSSFWorkbook workbook,
      XSSFDrawing drawing,
      String name,
      int col1,
      int row1,
      int col2,
      int row2)
      throws IOException {
    int storageId =
        workbook.addOlePackage(
            "payload".getBytes(StandardCharsets.UTF_8), "Payload", "payload.txt", "payload.txt");
    int pictureIndex = workbook.addPicture(PNG_PIXEL_BYTES, Workbook.PICTURE_TYPE_PNG);
    XSSFObjectData objectData =
        drawing.createObjectData(
            poiAnchor(drawing, col1, row1, col2, row2), storageId, pictureIndex);
    objectData.getCTShape().getNvSpPr().getCNvPr().setName(name);
    return objectData;
  }

  private static XSSFClientAnchor poiAnchor(
      XSSFDrawing drawing, int col1, int row1, int col2, int row2) {
    return drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
  }

  /** Minimal package-part base class for synthetic OPC edge-case fixtures. */
  private abstract static class TestPackagePart extends PackagePart {
    private final OPCPackage container;

    private TestPackagePart(String partName, String contentType) throws InvalidFormatException {
      this(OPCPackage.create(new ByteArrayOutputStream()), partName, contentType);
    }

    private TestPackagePart(OPCPackage container, String partName, String contentType)
        throws InvalidFormatException {
      super(container, PackagingURIHelper.createPartName(partName), contentType);
      this.container = container;
    }

    @Override
    protected OutputStream getOutputStreamImpl() {
      return new ByteArrayOutputStream();
    }

    @Override
    public boolean save(OutputStream outputStream) {
      return true;
    }

    @Override
    public boolean load(InputStream inputStream) {
      return true;
    }

    @Override
    public void close() {
      try {
        container.close();
      } catch (IOException exception) {
        throw new IllegalStateException("Failed to close test package", exception);
      }
    }

    @Override
    public void flush() {}
  }

  /** Synthetic package part whose relationship enumeration always fails. */
  private static final class InvalidRelationshipsPackagePart extends TestPackagePart {
    private InvalidRelationshipsPackagePart(String partName, String contentType)
        throws InvalidFormatException {
      super(partName, contentType);
    }

    @Override
    public PackageRelationshipCollection getRelationships() throws InvalidFormatException {
      throw new InvalidFormatException("broken relationships");
    }

    @Override
    protected InputStream getInputStreamImpl() {
      return new ByteArrayInputStream(new byte[0]);
    }
  }
}
