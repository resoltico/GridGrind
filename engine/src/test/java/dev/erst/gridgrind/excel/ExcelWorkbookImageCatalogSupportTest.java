package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.invoke.MethodHandles;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused regressions for workbook image-catalog synchronization around drawing media. */
class ExcelWorkbookImageCatalogSupportTest {
  private static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  @Test
  void synchronizePictureCatalogIncludesSignatureLinePreviewImages() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setPicture(pictureDefinition("OpsPicture", 1, 1, 4, 6));
      sheet.setSignatureLine(signatureDefinition("OpsSignature", 5, 1, 8, 6));

      assertEquals(
          List.of("/xl/media/image1.png"),
          ExcelWorkbookImageCatalogSupport.pictureCatalogPartNames(workbook.xssfWorkbook()));
      assertEquals(
          List.of("/xl/media/image1.png", "/xl/media/image2.png"),
          ExcelWorkbookImageCatalogSupport.packageImagePartNames(workbook.xssfWorkbook()));

      ExcelWorkbookImageCatalogSupport.synchronizePictureCatalog(workbook.xssfWorkbook());

      assertEquals(
          ExcelWorkbookImageCatalogSupport.packageImagePartNames(workbook.xssfWorkbook()),
          ExcelWorkbookImageCatalogSupport.pictureCatalogPartNames(workbook.xssfWorkbook()));
    }
  }

  @Test
  void embeddedObjectCreationSurvivesExistingSignaturePreviewMediaParts() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-drawing-media-catalog-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setPicture(pictureDefinition("OpsPicture", 1, 1, 4, 6));
      sheet.setSignatureLine(signatureDefinition("OpsSignature", 5, 1, 8, 6));

      assertDoesNotThrow(
          () ->
              sheet.setEmbeddedObject(
                  new ExcelEmbeddedObjectDefinition(
                      "OpsEmbed",
                      "Payload",
                      "payload.txt",
                      "payload.txt",
                      new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
                      ExcelPictureFormat.PNG,
                      new ExcelBinaryData(PNG_PIXEL_BYTES),
                      anchor(9, 1, 12, 6))));
      assertEquals(
          List.of("/xl/media/image1.png", "/xl/media/image2.png", "/xl/media/image3.png"),
          ExcelWorkbookImageCatalogSupport.packageImagePartNames(workbook.xssfWorkbook()));

      workbook.save(workbookPath);
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      assertEquals(
          List.of("/xl/media/image1.png", "/xl/media/image2.png", "/xl/media/image3.png"),
          ExcelWorkbookImageCatalogSupport.packageImagePartNames(workbook));
      assertEquals(
          List.of("OpsPicture", "OpsEmbed"),
          workbook.getSheet("Ops").getDrawingPatriarch().getShapes().stream()
              .map(XSSFShape::getShapeName)
              .toList());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      ExcelDrawingObjectPayload.EmbeddedObject payload =
          assertInstanceOf(
              ExcelDrawingObjectPayload.EmbeddedObject.class,
              workbook.sheet("Ops").drawingObjectPayload("OpsEmbed"));
      assertArrayEquals("payload".getBytes(StandardCharsets.UTF_8), payload.data().bytes());
      assertEquals(
          List.of("OpsPicture", "OpsEmbed", "OpsSignature"),
          workbook.sheet("Ops").drawingObjects().stream()
              .map(ExcelDrawingObjectSnapshot::name)
              .toList());
    }
  }

  @Test
  void addPictureSurvivesNonContiguousExistingImagePartNumbers() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.setPicture(pictureDefinition("OpsPicture", 1, 1, 4, 6));
      sheet.setEmbeddedObject(
          new ExcelEmbeddedObjectDefinition(
              "OpsEmbed",
              "Payload",
              "payload.txt",
              "payload.txt",
              new ExcelBinaryData("payload".getBytes(StandardCharsets.UTF_8)),
              ExcelPictureFormat.PNG,
              new ExcelBinaryData(PNG_PIXEL_BYTES),
              anchor(5, 1, 8, 6)));

      sheet.deleteDrawingObject("OpsPicture");

      assertEquals(
          List.of("/xl/media/image2.png"),
          ExcelWorkbookImageCatalogSupport.packageImagePartNames(workbook.xssfWorkbook()));
      assertEquals(
          List.of("/xl/media/image2.png"),
          ExcelWorkbookImageCatalogSupport.pictureCatalogPartNames(workbook.xssfWorkbook()));

      assertDoesNotThrow(() -> sheet.setPicture(pictureDefinition("OpsPicture2", 9, 1, 12, 6)));
      assertEquals(
          List.of("/xl/media/image2.png", "/xl/media/image3.png"),
          ExcelWorkbookImageCatalogSupport.packageImagePartNames(workbook.xssfWorkbook()));
      assertEquals(
          List.of("OpsEmbed", "OpsPicture2"),
          sheet.drawingObjects().stream().map(ExcelDrawingObjectSnapshot::name).toList());
    }
  }

  @Test
  @SuppressWarnings("PMD.CloseResource")
  void imageCatalogOrdersCustomMediaPartsAndFiltersInvalidContentTypes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      OPCPackage pkg = workbook.getPackage();
      pkg.createPart(PackagingURIHelper.createPartName("/xl/media/custom.png"), "image/png");
      pkg.createPart(PackagingURIHelper.createPartName("/xl/media/image2.jpeg"), "image/jpeg");
      pkg.createPart(PackagingURIHelper.createPartName("/xl/media/image2.png"), "image/png");
      pkg.createPart(
          PackagingURIHelper.createPartName("/xl/media/not-image.bin"), "application/octet-stream");

      assertEquals(
          List.of("/xl/media/image2.jpeg", "/xl/media/image2.png", "/xl/media/custom.png"),
          ExcelWorkbookImageCatalogSupport.packageImagePartNames(workbook));

      ExcelWorkbookImageCatalogSupport.synchronizePictureCatalog(workbook);

      assertEquals(
          List.of("/xl/media/image2.jpeg", "/xl/media/image2.png", "/xl/media/custom.png"),
          ExcelWorkbookImageCatalogSupport.pictureCatalogPartNames(workbook));
    }
  }

  @Test
  void imageCatalogSupportCoversLookupAndErrorGuards() {
    for (ExcelPictureFormat format : ExcelPictureFormat.values()) {
      POIXMLRelation relation =
          ExcelWorkbookImageCatalogSupport.pictureRelation(
              ExcelPicturePoiBridge.toPoiPictureType(format));
      assertNotNull(relation);
      assertEquals(format.defaultContentType(), relation.getContentType());
      assertEquals(relation, ExcelWorkbookImageCatalogSupport.pictureRelation(format));
    }
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelWorkbookImageCatalogSupport.pictureRelation(Integer.MIN_VALUE));
    assertEquals(3, ExcelWorkbookImageCatalogSupport.requireAllocatedPartNumber(3));
    IllegalStateException missingPartNumber =
        assertThrows(
            IllegalStateException.class,
            () -> ExcelWorkbookImageCatalogSupport.requireAllocatedPartNumber(-1));
    assertEquals("Failed to allocate a workbook image part", missingPartNumber.getMessage());
    IllegalStateException invalidImageParts =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelWorkbookImageCatalogSupport.imageParts(
                    () -> {
                      throw new InvalidFormatException("broken");
                    }));
    assertEquals("Failed to inspect workbook package media parts", invalidImageParts.getMessage());
    assertEquals("broken", invalidImageParts.getCause().getMessage());
    assertThrows(
        IllegalStateException.class,
        () -> ExcelWorkbookImageCatalogSupport.requirePicturesField(MethodHandles.publicLookup()));
    assertThrows(
        IllegalStateException.class,
        () ->
            ExcelWorkbookImageCatalogSupport.requirePictureConstructor(
                MethodHandles.publicLookup()));
  }

  @Test
  void writePictureBytesWrapsIoFailures() throws Exception {
    try (OPCPackage pkg = OPCPackage.create(new ByteArrayOutputStream())) {
      ThrowingPackagePart part =
          new ThrowingPackagePart(pkg, PackagingURIHelper.createPartName("/xl/media/image1.png"));

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class,
              () -> ExcelWorkbookImageCatalogSupport.writePictureBytes(part, PNG_PIXEL_BYTES));

      assertEquals("Failed to write workbook image bytes", failure.getMessage());
      assertEquals("broken", failure.getCause().getMessage());
    }
  }

  private static ExcelPictureDefinition pictureDefinition(
      String objectName, int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelPictureDefinition(
        objectName,
        new ExcelBinaryData(PNG_PIXEL_BYTES),
        ExcelPictureFormat.PNG,
        anchor(fromColumn, fromRow, toColumn, toRow),
        "preview");
  }

  private static ExcelSignatureLineDefinition signatureDefinition(
      String objectName, int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelSignatureLineDefinition(
        objectName,
        anchor(fromColumn, fromRow, toColumn, toRow),
        false,
        "Review before signing.",
        "Ada Lovelace",
        "Finance",
        "ada@example.com",
        null,
        "invalid",
        ExcelPictureFormat.PNG,
        new ExcelBinaryData(PNG_PIXEL_BYTES));
  }

  private static ExcelDrawingAnchor.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow, 0, 0),
        new ExcelDrawingMarker(toColumn, toRow, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
  }

  /** Test double that surfaces output-stream failures without depending on workbook package I/O. */
  private static final class ThrowingPackagePart extends PackagePart {
    private ThrowingPackagePart(OPCPackage container, PackagePartName partName)
        throws InvalidFormatException {
      super(container, partName, "image/png");
    }

    @Override
    protected InputStream getInputStreamImpl() {
      return InputStream.nullInputStream();
    }

    @Override
    protected OutputStream getOutputStreamImpl() throws IOException {
      throw new IOException("broken");
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
    public void close() {}

    @Override
    public void flush() {}
  }
}
