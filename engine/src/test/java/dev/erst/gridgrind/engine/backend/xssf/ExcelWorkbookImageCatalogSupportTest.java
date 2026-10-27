package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.engine.drawing.DrawingAdapter;
import dev.erst.gridgrind.engine.poi.ImageCatalog;
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
import java.util.Optional;
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
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");

  @Test
  void synchronizePictureCatalogIncludesSignatureLinePreviewImages() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.drawings().setPicture(pictureDefinition("OpsPicture", 1, 1, 4, 6));
      sheet.drawings().setSignatureLine(signatureDefinition("OpsSignature", 5, 1, 8, 6));

      assertEquals(
          List.of("/xl/media/image1.png"),
          ImageCatalog.pictureCatalogPartNames(workbook.xssfWorkbook()));
      assertEquals(
          List.of("/xl/media/image1.png", "/xl/media/image2.png"),
          ImageCatalog.packageImagePartNames(workbook.xssfWorkbook()));

      ImageCatalog.synchronizePictureCatalog(workbook.xssfWorkbook());

      assertEquals(
          ImageCatalog.packageImagePartNames(workbook.xssfWorkbook()),
          ImageCatalog.pictureCatalogPartNames(workbook.xssfWorkbook()));
    }
  }

  @Test
  void embeddedObjectCreationSurvivesExistingSignaturePreviewMediaParts() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-drawing-media-catalog-");

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.drawings().setPicture(pictureDefinition("OpsPicture", 1, 1, 4, 6));
      sheet.drawings().setSignatureLine(signatureDefinition("OpsSignature", 5, 1, 8, 6));

      assertDoesNotThrow(
          () ->
              sheet
                  .drawings()
                  .setEmbeddedObject(
                      new ResolvedEmbeddedObject(
                          "OpsEmbed",
                          "Payload",
                          "payload.txt",
                          "payload.txt",
                          new BinaryContent("payload".getBytes(StandardCharsets.UTF_8)),
                          ExcelPictureFormat.PNG,
                          new BinaryContent(PNG_PIXEL_BYTES),
                          anchor(9, 1, 12, 6))));
      assertEquals(
          List.of("/xl/media/image1.png", "/xl/media/image2.png", "/xl/media/image3.png"),
          ImageCatalog.packageImagePartNames(workbook.xssfWorkbook()));

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(workbookPath))) {
      assertEquals(
          List.of("/xl/media/image1.png", "/xl/media/image2.png", "/xl/media/image3.png"),
          ImageCatalog.packageImagePartNames(workbook));
      assertEquals(
          List.of("OpsPicture", "OpsEmbed"),
          workbook.getSheet("Ops").getDrawingPatriarch().getShapes().stream()
              .map(XSSFShape::getShapeName)
              .toList());
    }

    try (XssfWorkbook workbook =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      DrawingObjectPayloadReport.EmbeddedObject payload =
          assertInstanceOf(
              DrawingObjectPayloadReport.EmbeddedObject.class,
              workbook.sheet("Ops").drawings().drawingObjectPayload("OpsEmbed"));
      assertArrayEquals(
          "payload".getBytes(StandardCharsets.UTF_8),
          java.util.Base64.getDecoder().decode(payload.base64Data()));
      assertEquals(
          List.of("OpsPicture", "OpsEmbed", "OpsSignature"),
          workbook.sheet("Ops").drawings().drawingObjects().stream()
              .map(DrawingObjectReport::name)
              .toList());
    }
  }

  @Test
  void addPictureSurvivesNonContiguousExistingImagePartNumbers() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Ops");
      sheet.drawings().setPicture(pictureDefinition("OpsPicture", 1, 1, 4, 6));
      sheet
          .drawings()
          .setEmbeddedObject(
              new ResolvedEmbeddedObject(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new BinaryContent("payload".getBytes(StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new BinaryContent(PNG_PIXEL_BYTES),
                  anchor(5, 1, 8, 6)));

      sheet.drawings().deleteDrawingObject("OpsPicture");

      assertEquals(
          List.of("/xl/media/image2.png"),
          ImageCatalog.packageImagePartNames(workbook.xssfWorkbook()));
      assertEquals(
          List.of("/xl/media/image2.png"),
          ImageCatalog.pictureCatalogPartNames(workbook.xssfWorkbook()));

      assertDoesNotThrow(
          () -> sheet.drawings().setPicture(pictureDefinition("OpsPicture2", 9, 1, 12, 6)));
      assertEquals(
          List.of("/xl/media/image2.png", "/xl/media/image3.png"),
          ImageCatalog.packageImagePartNames(workbook.xssfWorkbook()));
      assertEquals(
          List.of("OpsEmbed", "OpsPicture2"),
          sheet.drawings().drawingObjects().stream().map(DrawingObjectReport::name).toList());
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
          ImageCatalog.packageImagePartNames(workbook));

      ImageCatalog.synchronizePictureCatalog(workbook);

      assertEquals(
          List.of("/xl/media/image2.jpeg", "/xl/media/image2.png", "/xl/media/custom.png"),
          ImageCatalog.pictureCatalogPartNames(workbook));
    }
  }

  @Test
  void imageCatalogSupportCoversLookupAndErrorGuards() {
    for (ExcelPictureFormat format : ExcelPictureFormat.values()) {
      POIXMLRelation relation =
          ImageCatalog.pictureRelation(DrawingAdapter.toPoiPictureType(format));
      assertNotNull(relation);
      assertEquals(format.defaultContentType(), relation.getContentType());
      assertEquals(relation, ImageCatalog.pictureRelation(format));
    }
    assertThrows(
        IllegalArgumentException.class, () -> ImageCatalog.pictureRelation(Integer.MIN_VALUE));
    assertEquals(3, ImageCatalog.requireAllocatedPartNumber(3));
    IllegalStateException missingPartNumber =
        assertThrows(
            IllegalStateException.class, () -> ImageCatalog.requireAllocatedPartNumber(-1));
    assertEquals("Failed to allocate a workbook image part", missingPartNumber.getMessage());
    IllegalStateException invalidImageParts =
        assertThrows(
            IllegalStateException.class,
            () ->
                ImageCatalog.imageParts(
                    () -> {
                      throw new InvalidFormatException("broken");
                    }));
    assertEquals("Failed to inspect workbook package media parts", invalidImageParts.getMessage());
    assertEquals("broken", invalidImageParts.getCause().getMessage());
    assertThrows(
        IllegalStateException.class,
        () -> ImageCatalog.requirePicturesField(MethodHandles.publicLookup()));
    IllegalStateException picturesFieldFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ImageCatalog.requirePicturesField(MethodHandles.publicLookup()));
    IllegalStateException pictureConstructorFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ImageCatalog.requirePictureConstructor(MethodHandles.publicLookup()));
    assertTrue(
        picturesFieldFailure.getMessage().contains("Apache POI private contract unavailable"));
    assertTrue(
        picturesFieldFailure
            .getMessage()
            .contains(ImageCatalog.PICTURES_FIELD_CONTRACT.affectedSurface()));
    assertTrue(
        pictureConstructorFailure
            .getMessage()
            .contains(ImageCatalog.PICTURE_CONSTRUCTOR_CONTRACT.affectedSurface()));
  }

  @Test
  void writePictureBytesWrapsIoFailures() throws Exception {
    try (OPCPackage pkg = OPCPackage.create(new ByteArrayOutputStream())) {
      ThrowingPackagePart part =
          new ThrowingPackagePart(pkg, PackagingURIHelper.createPartName("/xl/media/image1.png"));

      IllegalStateException failure =
          assertThrows(
              IllegalStateException.class,
              () -> ImageCatalog.writePictureBytes(part, PNG_PIXEL_BYTES));

      assertEquals("Failed to write workbook image bytes", failure.getMessage());
      assertEquals("broken", failure.getCause().getMessage());
    }
  }

  private static ResolvedPicture pictureDefinition(
      String objectName, int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ResolvedPicture(
        objectName,
        new BinaryContent(PNG_PIXEL_BYTES),
        ExcelPictureFormat.PNG,
        anchor(fromColumn, fromRow, toColumn, toRow),
        Optional.of("preview"));
  }

  private static ResolvedSignatureLine signatureDefinition(
      String objectName, int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ResolvedSignatureLine(
        objectName,
        anchor(fromColumn, fromRow, toColumn, toRow),
        false,
        Optional.of("Review before signing."),
        Optional.of("Ada Lovelace"),
        Optional.of("Finance"),
        Optional.of("ada@example.com"),
        Optional.empty(),
        Optional.of("invalid"),
        Optional.of(ExcelPictureFormat.PNG),
        Optional.of(new BinaryContent(PNG_PIXEL_BYTES)));
  }

  private static DrawingAnchorSpec.TwoCell anchor(
      int fromColumn, int fromRow, int toColumn, int toRow) {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(fromColumn, fromRow, 0, 0),
        new DrawingMarkerSpec(toColumn, toRow, 0, 0),
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
