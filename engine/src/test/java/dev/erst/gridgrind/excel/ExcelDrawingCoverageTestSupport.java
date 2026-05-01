package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Base64;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFConnector;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.STEditAs;

/**
 * Shared helpers for drawing coverage slices.
 *
 * <p>This support surface intentionally centralizes terse nested fixtures and checked-exception
 * helpers so the individual coverage slices can stay focused on behavioral theory.
 */
@SuppressWarnings({
  "PMD.CommentRequired",
  "PMD.SignatureDeclareThrowsException",
  "PMD.UseUtilityClass"
})
class ExcelDrawingCoverageTestSupport {
  static final byte[] PNG_PIXEL_BYTES =
      Base64.getDecoder()
          .decode(
              "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

  static XSSFPicture createPicture(
      XSSFWorkbook workbook,
      XSSFDrawing drawing,
      String name,
      int col1,
      int row1,
      int col2,
      int row2) {
    int pictureIndex = workbook.addPicture(PNG_PIXEL_BYTES, Workbook.PICTURE_TYPE_PNG);
    XSSFPicture picture =
        drawing.createPicture(poiAnchor(drawing, col1, row1, col2, row2), pictureIndex);
    picture.getCTPicture().getNvPicPr().getCNvPr().setName(name);
    return picture;
  }

  static XSSFSimpleShape createSimpleShape(
      XSSFDrawing drawing, String name, int col1, int row1, int col2, int row2) {
    XSSFSimpleShape shape = drawing.createSimpleShape(poiAnchor(drawing, col1, row1, col2, row2));
    shape.getCTShape().getNvSpPr().getCNvPr().setName(name);
    return shape;
  }

  static XSSFConnector createConnector(
      XSSFDrawing drawing, String name, int col1, int row1, int col2, int row2) {
    XSSFConnector connector = drawing.createConnector(poiAnchor(drawing, col1, row1, col2, row2));
    connector.getCTConnector().getNvCxnSpPr().getCNvPr().setName(name);
    return connector;
  }

  static XSSFObjectData createEmbeddedObject(
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

  static XSSFClientAnchor poiAnchor(XSSFDrawing drawing, int col1, int row1, int col2, int row2) {
    return drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
  }

  static ExcelDrawingAnchor.TwoCell anchor(int fromColumn, int fromRow, int toColumn, int toRow) {
    return new ExcelDrawingAnchor.TwoCell(
        new ExcelDrawingMarker(fromColumn, fromRow),
        new ExcelDrawingMarker(toColumn, toRow),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  static CTShape oneCellShape() {
    CTOneCellAnchor anchor = CTOneCellAnchor.Factory.newInstance();
    populateOneCellAnchor(anchor);
    return anchor.getSp();
  }

  static CTShape absoluteShape() {
    CTAbsoluteAnchor anchor = CTAbsoluteAnchor.Factory.newInstance();
    populateAbsoluteAnchor(anchor);
    return anchor.getSp();
  }

  static CTShape unsupportedParentShape() {
    return org.openxmlformats
        .schemas
        .drawingml
        .x2006
        .spreadsheetDrawing
        .CTGroupShape
        .Factory
        .newInstance()
        .addNewSp();
  }

  static void populateTwoCellAnchor(CTTwoCellAnchor anchor) {
    CTMarker from = anchor.addNewFrom();
    from.setCol(1);
    from.setRow(2);
    from.setColOff(3L);
    from.setRowOff(4L);
    CTMarker to = anchor.addNewTo();
    to.setCol(5);
    to.setRow(6);
    to.setColOff(7L);
    to.setRowOff(8L);
    anchor.addNewSp();
  }

  static void populateOneCellAnchor(CTOneCellAnchor anchor) {
    CTMarker from = anchor.addNewFrom();
    from.setCol(1);
    from.setRow(2);
    from.setColOff(3L);
    from.setRowOff(4L);
    anchor.addNewExt().setCx(5L);
    anchor.getExt().setCy(6L);
    anchor.addNewSp();
  }

  static void populateAbsoluteAnchor(CTAbsoluteAnchor anchor) {
    anchor.addNewPos().setX(7L);
    anchor.getPos().setY(8L);
    anchor.addNewExt().setCx(9L);
    anchor.getExt().setCy(10L);
    anchor.addNewSp();
  }

  static byte[] ole2StorageBytes() throws IOException {
    try (var filesystem = new org.apache.poi.poifs.filesystem.POIFSFileSystem();
        var output = new ByteArrayOutputStream()) {
      filesystem.createDirectory("Root");
      filesystem.writeFilesystem(output);
      return output.toByteArray();
    }
  }

  static <T> T invoke(Object target, String name, Class<T> returnType, Object... args)
      throws Exception {
    return returnType.cast(dispatch(target, name, args));
  }

  static void invokeVoid(Object target, String name, Object... args) throws Exception {
    dispatch(target, name, args);
  }

  static Object dispatch(Object target, String name, Object... args) throws Exception {
    if (target instanceof ExcelDrawingController controller) {
      return switch (name) {
        case "behavior" -> controller.behavior((STEditAs.Enum) args[0]);
        case "binary" -> controller.binary((byte[]) args[0], (String) args[1]);
        case "cleanupPackagePartIfUnused" -> {
          controller.cleanupPackagePartIfUnused(
              (OPCPackage) args[0], (org.apache.poi.openxml4j.opc.PackagePartName) args[1]);
          yield null;
        }
        case "cleanupWorkbookImagePartIfUnused" -> {
          controller.cleanupWorkbookImagePartIfUnused(
              (XSSFWorkbook) args[0], (org.apache.poi.openxml4j.opc.PackagePartName) args[1]);
          yield null;
        }
        case "defaultName" -> controller.defaultName((XSSFShape) args[0]);
        case "firstNonBlank" -> controller.firstNonBlank((String) args[0], (String) args[1]);
        case "imagePartUsed" ->
            controller.imagePartUsed(
                (XSSFWorkbook) args[0], (org.apache.poi.openxml4j.opc.PackagePartName) args[1]);
        case "looksLikeOle2Storage" -> controller.looksLikeOle2Storage((byte[]) args[0]);
        case "parentAnchor" -> controller.parentAnchor((org.apache.xmlbeans.XmlObject) args[0]);
        case "partBytes" -> controller.partBytes((PackagePart) args[0]);
        case "partFileName" -> controller.partFileName((PackagePart) args[0]);
        case "previewDrawingRelationId" ->
            controller.previewDrawingRelationId((XSSFObjectData) args[0]);
        case "previewImagePart" -> controller.previewImagePart((XSSFObjectData) args[0]);
        case "previewSheetRelationId" ->
            controller.previewSheetRelationId(
                (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject) args[0]);
        case "rasterDimensions" -> ExcelDrawingController.rasterDimensions((byte[]) args[0]);
        case "relatedInternalPart" ->
            controller.relatedInternalPart((PackagePart) args[0], (String) args[1]);
        case "removeAbsoluteAnchor" -> {
          controller.removeAbsoluteAnchor((CTDrawing) args[0], (CTAbsoluteAnchor) args[1]);
          yield null;
        }
        case "removeOleObject" -> {
          controller.removeOleObject(
              (XSSFSheet) args[0],
              (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOleObject) args[1]);
          yield null;
        }
        case "removeOneCellAnchor" -> {
          controller.removeOneCellAnchor((CTDrawing) args[0], (CTOneCellAnchor) args[1]);
          yield null;
        }
        case "removeParentAnchor" -> {
          controller.removeParentAnchor(
              (XSSFDrawing) args[0], (org.apache.xmlbeans.XmlObject) args[1]);
          yield null;
        }
        case "removeRelationshipsToPart" -> {
          controller.removeRelationshipsToPart(
              (PackagePart) args[0], (org.apache.poi.openxml4j.opc.PackagePartName) args[1]);
          yield null;
        }
        case "removeTwoCellAnchor" -> {
          controller.removeTwoCellAnchor((CTDrawing) args[0], (CTTwoCellAnchor) args[1]);
          yield null;
        }
        case "requireNonBlank" -> controller.requireNonBlank((String) args[0], (String) args[1]);
        case "resolvedName" -> controller.resolvedName((XSSFShape) args[0]);
        case "sha256" -> controller.sha256((byte[]) args[0]);
        case "shapeType" -> controller.shapeType((String) args[0]);
        case "shapeXml" -> controller.shapeXml((XSSFShape) args[0]);
        case "snapshot" ->
            ExcelDrawingSnapshotSupport.snapshot((XSSFDrawing) args[0], (XSSFShape) args[1]);
        case "snapshotAnchor" -> controller.snapshotAnchor((org.apache.xmlbeans.XmlObject) args[0]);
        case "snapshotShape" -> controller.snapshotShape((XSSFSimpleShape) args[0]);
        case "toPoiBehavior" -> controller.toPoiBehavior((ExcelDrawingAnchorBehavior) args[0]);
        case "toPoiEditAs" -> controller.toPoiEditAs((ExcelDrawingAnchorBehavior) args[0]);
        case "updateAnchorInPlace" -> {
          controller.updateAnchorInPlace(
              (XSSFSheet) args[0],
              (String) args[1],
              (org.apache.xmlbeans.XmlObject) args[2],
              (ExcelDrawingAnchor.TwoCell) args[3]);
          yield null;
        }
        default -> throw new IllegalArgumentException("Unsupported helper invocation: " + name);
      };
    }
    if (target instanceof ExcelDrawingController.RasterDimensions dimensions) {
      return switch (name) {
        case "widthPixels" -> dimensions.widthPixels();
        case "heightPixels" -> dimensions.heightPixels();
        default -> throw new IllegalArgumentException("Unsupported helper invocation: " + name);
      };
    }
    if (target instanceof ExcelDrawingSnapshotSupport.RasterDimensions dimensions) {
      return switch (name) {
        case "widthPixels" -> dimensions.widthPixels();
        case "heightPixels" -> dimensions.heightPixels();
        default -> throw new IllegalArgumentException("Unsupported helper invocation: " + name);
      };
    }
    throw new IllegalArgumentException(
        "Unsupported helper target: " + target.getClass().getName() + "#" + name);
  }

  static <T extends Throwable> T assertInvocationFailure(Class<T> type, ThrowingRunnable runnable) {
    return assertThrows(type, runnable::run);
  }

  @FunctionalInterface
  interface ThrowingRunnable {
    void run() throws Exception;
  }

  static final class DefinedNameStub implements org.apache.poi.ss.usermodel.Name {
    final String name;
    final String refersToFormula;
    final int sheetIndex;

    DefinedNameStub(String name, String refersToFormula, int sheetIndex) {
      this.name = name;
      this.refersToFormula = refersToFormula;
      this.sheetIndex = sheetIndex;
    }

    @Override
    public String getSheetName() {
      return sheetIndex < 0 ? null : "Ops";
    }

    @Override
    public String getNameName() {
      return name;
    }

    @Override
    public void setNameName(String name) {
      throw new UnsupportedOperationException();
    }

    @Override
    public String getRefersToFormula() {
      return refersToFormula;
    }

    @Override
    public void setRefersToFormula(String formulaText) {
      throw new UnsupportedOperationException();
    }

    @Override
    public boolean isFunctionName() {
      return false;
    }

    @Override
    public boolean isDeleted() {
      return false;
    }

    @Override
    public boolean isHidden() {
      return false;
    }

    @Override
    public void setSheetIndex(int index) {
      throw new UnsupportedOperationException();
    }

    @Override
    public int getSheetIndex() {
      return sheetIndex;
    }

    @Override
    public String getComment() {
      return "";
    }

    @Override
    public void setComment(String comment) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setFunction(boolean value) {
      throw new UnsupportedOperationException();
    }
  }

  static final class UnsupportedShape extends XSSFShape {
    @Override
    public String getShapeName() {
      return "Unsupported";
    }

    @Override
    protected CTShapeProperties getShapeProperties() {
      return null;
    }
  }

  abstract static class TestPackagePart extends PackagePart {
    final OPCPackage container;

    TestPackagePart(String partName, String contentType) throws InvalidFormatException {
      this(OPCPackage.create(new ByteArrayOutputStream()), partName, contentType);
    }

    TestPackagePart(OPCPackage container, String partName, String contentType)
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

  static final class FixedBytesPackagePart extends TestPackagePart {
    final byte[] bytes;

    FixedBytesPackagePart(String partName, String contentType, byte[] bytes)
        throws InvalidFormatException {
      super(partName, contentType);
      this.bytes = bytes.clone();
    }

    @Override
    protected InputStream getInputStreamImpl() {
      return new ByteArrayInputStream(bytes);
    }
  }

  static final class FailingPackagePart extends TestPackagePart {
    FailingPackagePart(String partName, String contentType) throws InvalidFormatException {
      super(partName, contentType);
    }

    @Override
    protected InputStream getInputStreamImpl() throws IOException {
      throw new IOException("broken");
    }
  }

  static final class InvalidTargetPackagePart extends TestPackagePart {
    InvalidTargetPackagePart(String partName, String contentType) throws InvalidFormatException {
      super(partName, contentType);
    }

    @Override
    public PackageRelationship getRelationship(String relationshipId) {
      return new PackageRelationship(
          getPackage(),
          this,
          java.net.URI.create(".."),
          TargetMode.INTERNAL,
          "urn:test",
          relationshipId);
    }

    @Override
    protected InputStream getInputStreamImpl() {
      return new ByteArrayInputStream(new byte[0]);
    }
  }

  static final class InvalidRelationshipsPackagePart extends TestPackagePart {
    InvalidRelationshipsPackagePart(String partName, String contentType)
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

  static final class RelationshipPartPackagePart extends TestPackagePart {
    RelationshipPartPackagePart(String partName, String contentType) throws InvalidFormatException {
      super(partName, contentType);
    }

    @Override
    public boolean isRelationshipPart() {
      return true;
    }

    @Override
    protected InputStream getInputStreamImpl() {
      return new ByteArrayInputStream(new byte[0]);
    }
  }
}
