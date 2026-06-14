package dev.erst.gridgrind.excel.drawing;

import com.microsoft.schemas.office.excel.CTClientData;
import com.microsoft.schemas.office.office.CTSignatureLine;
import com.microsoft.schemas.vml.CTShape;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSignatureLine;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STTrueFalse;

/** Reads and mutates VML-backed Excel signature lines. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelSignatureLineController {
  private static final Pattern VML_ANCHOR_SEPARATOR = Pattern.compile("\\s*,\\s*");

  public List<ExcelSignatureLineSnapshot> signatureLines(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return List.of();
    }

    List<ExcelSignatureLineSnapshot> snapshots = new ArrayList<>();
    int index = 1;
    for (CTShape shape : ExcelSignatureLineSnapshotSupport.signatureShapes(vmlDrawing)) {
      int currentIndex = index;
      index++;
      snapshots.add(snapshot(vmlDrawing, shape, currentIndex));
    }
    return List.copyOf(snapshots);
  }

  public void setSignatureLine(XSSFSheet sheet, ExcelSignatureLineDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteIfPresent(sheet, definition.name());

    XSSFSignatureLine signatureLine = configuredSignatureLine(definition);
    signatureLine.add(sheet, toPoiAnchor(definition.anchor()));
    applyNameAndCommentMetadata(signatureLine.getSignatureShape(), definition);
  }

  public boolean updateAnchorIfPresent(
      XSSFSheet sheet, String objectName, ExcelDrawingAnchor.TwoCell anchor) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");
    Objects.requireNonNull(anchor, "anchor must not be null");

    Optional<LocatedSignatureLine> located = find(sheet, objectName);
    if (located.isEmpty()) {
      return false;
    }
    CTClientData clientData = requiredClientData(sheet, located.orElseThrow().shape(), objectName);
    clientData.setAnchorArray(0, anchorString(anchor));
    return true;
  }

  public boolean deleteIfPresent(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");

    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return false;
    }

    Optional<LocatedSignatureLine> located = find(sheet, objectName);
    if (located.isEmpty()) {
      return false;
    }

    LocatedSignatureLine resolved = located.orElseThrow();
    PackagePart previewPart = previewPart(vmlDrawing, resolved.shape());
    String previewRelationId = previewRelationId(resolved.shape());
    if (previewRelationId != null) {
      vmlDrawing.getPackagePart().removeRelationship(previewRelationId);
    }
    resolved.cursor().removeXml();
    if (previewPart != null) {
      ExcelDrawingBinarySupport.cleanupWorkbookImagePartIfUnused(
          sheet.getWorkbook(), previewPart.getPartName());
    }
    return true;
  }

  public boolean hasNamedSignatureLine(XSSFSheet sheet, String objectName) {
    return find(sheet, objectName).isPresent();
  }

  private ExcelSignatureLineSnapshot snapshot(XSSFVMLDrawing vmlDrawing, CTShape shape, int index) {
    String objectName = ExcelSignatureLineSnapshotSupport.resolvedName(shape, index);
    PackagePart previewPart = previewPart(vmlDrawing, shape);
    ExcelPictureFormat previewFormat =
        previewPart == null
            ? null
            : ExcelPictureFormat.fromContentType(previewPart.getContentType());
    byte[] previewBytes =
        previewPart == null ? null : ExcelDrawingBinarySupport.partBytes(previewPart);
    ExcelDrawingSnapshotSupport.RasterDimensions previewDimensions =
        previewBytes == null
            ? ExcelDrawingSnapshotSupport.RasterDimensions.none()
            : ExcelDrawingSnapshotSupport.rasterDimensions(previewBytes);
    CTSignatureLine signatureLine = shape.getSignaturelineArray(0);
    return new ExcelSignatureLineSnapshot(
        objectName,
        anchor(shape, objectName),
        ExcelSignatureLineSnapshotSupport.signatureLineSetup(
            signatureLine, ExcelSignatureLineSnapshotSupport.signingInstructions(signatureLine)),
        ExcelSignatureLineSnapshotSupport.signatureLinePreview(
            previewFormat, previewPart, previewBytes, previewDimensions));
  }

  private Optional<LocatedSignatureLine> find(XSSFSheet sheet, String objectName) {
    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return Optional.empty();
    }

    CTShape matchedShape = null;
    XmlCursor matchedCursor = null;
    int matchedIndex = -1;
    int index = 1;
    try (XmlCursor cursor = vmlDrawing.getDocument().getXml().newCursor()) {
      for (boolean found = cursor.toFirstChild(); found; found = cursor.toNextSibling()) {
        XmlObject object = cursor.getObject();
        if (!(object instanceof CTShape shape) || shape.sizeOfSignaturelineArray() == 0) {
          continue;
        }
        if (!ExcelSignatureLineSnapshotSupport.resolvedName(shape, index).equals(objectName)) {
          index++;
          continue;
        }
        if (matchedShape != null) {
          throw new IllegalArgumentException(
              "Multiple signature lines named '"
                  + objectName
                  + "' exist on sheet '"
                  + sheet.getSheetName()
                  + "'");
        }
        matchedShape = shape;
        matchedCursor = cursor.newCursor();
        matchedIndex = index;
        index++;
      }
    }
    return matchedShape == null
        ? Optional.empty()
        : Optional.of(
            new LocatedSignatureLine(
                matchedShape,
                Objects.requireNonNull(matchedCursor, "matchedCursor must not be null"),
                matchedIndex));
  }

  private static XSSFSignatureLine configuredSignatureLine(
      ExcelSignatureLineDefinition definition) {
    XSSFSignatureLine signatureLine = new XSSFSignatureLine();
    applyOptionalMetadata(signatureLine, definition);
    applyPlainSignature(signatureLine, definition);
    return signatureLine;
  }

  private static void applyOptionalMetadata(
      XSSFSignatureLine signatureLine, ExcelSignatureLineDefinition definition) {
    applyIfPresent(definition.signingInstructions(), signatureLine::setSigningInstructions);
    applyIfPresent(definition.suggestedSigner(), signatureLine::setSuggestedSigner);
    applyIfPresent(definition.suggestedSigner2(), signatureLine::setSuggestedSigner2);
    applyIfPresent(definition.suggestedSignerEmail(), signatureLine::setSuggestedSignerEmail);
    applyIfPresent(definition.caption(), signatureLine::setCaption);
    applyIfPresent(definition.invalidStamp(), signatureLine::setInvalidStamp);
  }

  private static void applyPlainSignature(
      XSSFSignatureLine signatureLine, ExcelSignatureLineDefinition definition) {
    if (definition.plainSignature().isEmpty()) {
      return;
    }
    signatureLine.setPlainSignature(definition.plainSignature().orElseThrow().bytes());
    signatureLine.setContentType(
        definition.plainSignatureFormat().orElseThrow().defaultContentType());
  }

  private static void applyNameAndCommentMetadata(
      CTShape shape, ExcelSignatureLineDefinition definition) {
    shape.setAlt(definition.name());
    applyImageTitle(shape, definition.name());
    shape
        .getSignaturelineArray(0)
        .setAllowcomments(definition.allowComments() ? STTrueFalse.T : STTrueFalse.F);
  }

  public static void applyImageTitle(CTShape shape, String name) {
    if (shape.sizeOfImagedataArray() == 0) {
      return;
    }
    shape.getImagedataArray(0).setTitle(name);
  }

  private static void applyIfPresent(
      @Nullable String value, java.util.function.Consumer<String> consumer) {
    if (value != null) {
      consumer.accept(value);
    }
  }

  public static boolean usesImagePart(
      XSSFSheet sheet, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(imagePartName, "imagePartName must not be null");
    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return false;
    }
    for (CTShape shape : ExcelSignatureLineSnapshotSupport.signatureShapes(vmlDrawing)) {
      PackagePart previewPart = previewPart(vmlDrawing, shape);
      if (previewPart != null && previewPart.getPartName().equals(imagePartName)) {
        return true;
      }
    }
    return false;
  }

  private static @Nullable PackagePart previewPart(XSSFVMLDrawing vmlDrawing, CTShape shape) {
    String previewRelationId = previewRelationId(shape);
    return previewRelationId == null
        ? null
        : ExcelDrawingBinarySupport.relatedInternalPart(
                vmlDrawing.getPackagePart(), previewRelationId)
            .orElse(null);
  }

  private static @Nullable String previewRelationId(CTShape shape) {
    return shape.sizeOfImagedataArray() == 0
        ? null
        : ExcelDrawingBinarySupport.blankAsOptional(shape.getImagedataArray(0).getRelid())
            .orElse(null);
  }

  private static ExcelDrawingAnchor.TwoCell anchor(CTShape shape, String objectName) {
    String value = requiredClientData(null, shape, objectName).getAnchorArray(0);
    String[] tokens = VML_ANCHOR_SEPARATOR.split(value, -1);
    if (tokens.length != 8) {
      throw new IllegalStateException(
          "Signature line '" + objectName + "' is backed by an invalid VML anchor: " + value);
    }
    try {
      return new ExcelDrawingAnchor.TwoCell(
          new ExcelDrawingMarker(
              Integer.parseInt(tokens[0]),
              Integer.parseInt(tokens[2]),
              Integer.parseInt(tokens[1]),
              Integer.parseInt(tokens[3])),
          new ExcelDrawingMarker(
              Integer.parseInt(tokens[4]),
              Integer.parseInt(tokens[6]),
              Integer.parseInt(tokens[5]),
              Integer.parseInt(tokens[7])),
          ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    } catch (NumberFormatException exception) {
      throw new IllegalStateException(
          "Signature line '" + objectName + "' is backed by an invalid VML anchor: " + value,
          exception);
    }
  }

  private static CTClientData requiredClientData(
      @Nullable XSSFSheet sheet, CTShape shape, String objectName) {
    if (shape.sizeOfClientDataArray() == 0) {
      throw new IllegalStateException(missingClientDataMessage(sheet, objectName));
    }
    CTClientData clientData = shape.getClientDataArray(0);
    if (clientData.sizeOfAnchorArray() == 0) {
      throw new IllegalStateException(missingAnchorMessage(sheet, objectName));
    }
    return clientData;
  }

  private static String anchorString(ExcelDrawingAnchor.TwoCell anchor) {
    return anchor.from().columnIndex()
        + ", "
        + anchor.from().dx()
        + ", "
        + anchor.from().rowIndex()
        + ", "
        + anchor.from().dy()
        + ", "
        + anchor.to().columnIndex()
        + ", "
        + anchor.to().dx()
        + ", "
        + anchor.to().rowIndex()
        + ", "
        + anchor.to().dy();
  }

  private static XSSFClientAnchor toPoiAnchor(ExcelDrawingAnchor.TwoCell anchor) {
    return new XSSFClientAnchor(
        anchor.from().dx(),
        anchor.from().dy(),
        anchor.to().dx(),
        anchor.to().dy(),
        anchor.from().columnIndex(),
        anchor.from().rowIndex(),
        anchor.to().columnIndex(),
        anchor.to().rowIndex());
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }

  private record LocatedSignatureLine(CTShape shape, XmlCursor cursor, int index) {}

  private static String missingClientDataMessage(@Nullable XSSFSheet sheet, String objectName) {
    return sheet == null
        ? "Signature line '" + objectName + "' is missing VML clientData"
        : "Signature line '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is missing VML clientData";
  }

  private static String missingAnchorMessage(@Nullable XSSFSheet sheet, String objectName) {
    return sheet == null
        ? "Signature line '" + objectName + "' is missing its VML anchor"
        : "Signature line '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is missing its VML anchor";
  }
}
