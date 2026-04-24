package dev.erst.gridgrind.excel;

import com.microsoft.schemas.office.excel.CTClientData;
import com.microsoft.schemas.office.office.CTSignatureLine;
import com.microsoft.schemas.vml.CTImageData;
import com.microsoft.schemas.vml.CTShape;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.regex.Pattern;
import javax.xml.namespace.QName;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSignatureLine;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STTrueFalse;

/** Reads and mutates VML-backed Excel signature lines. */
final class ExcelSignatureLineController {
  private static final String DEFAULT_ALT_TEXT = "Microsoft Office Signature Line...";
  private static final String MS_OFFICE_URN = "urn:schemas-microsoft-com:office:office";
  private static final Pattern VML_ANCHOR_SEPARATOR = Pattern.compile("\\s*,\\s*");
  private static final QName SIGNING_INSTRUCTIONS = new QName(MS_OFFICE_URN, "signinginstructions");

  List<ExcelSignatureLineSnapshot> signatureLines(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return List.of();
    }

    List<ExcelSignatureLineSnapshot> snapshots = new ArrayList<>();
    int index = 1;
    for (CTShape shape : signatureShapes(vmlDrawing)) {
      int currentIndex = index;
      index++;
      snapshots.add(snapshot(vmlDrawing, shape, currentIndex));
    }
    return List.copyOf(snapshots);
  }

  void setSignatureLine(XSSFSheet sheet, ExcelSignatureLineDefinition definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteIfPresent(sheet, definition.name());

    XSSFSignatureLine signatureLine = configuredSignatureLine(definition);
    signatureLine.add(sheet, toPoiAnchor(definition.anchor()));
    applyNameAndCommentMetadata(signatureLine.getSignatureShape(), definition);
  }

  boolean updateAnchorIfPresent(
      XSSFSheet sheet, String objectName, ExcelDrawingAnchor.TwoCell anchor) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");
    Objects.requireNonNull(anchor, "anchor must not be null");

    LocatedSignatureLine located = find(sheet, objectName);
    if (located == null) {
      return false;
    }
    CTClientData clientData = requiredClientData(sheet, located.shape(), objectName);
    clientData.setAnchorArray(0, anchorString(anchor));
    return true;
  }

  boolean deleteIfPresent(XSSFSheet sheet, String objectName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    requireNonBlank(objectName, "objectName");

    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return false;
    }

    LocatedSignatureLine located = find(sheet, objectName);
    if (located == null) {
      return false;
    }

    PackagePart previewPart = previewPart(vmlDrawing, located.shape());
    String previewRelationId = previewRelationId(located.shape());
    if (previewRelationId != null) {
      vmlDrawing.getPackagePart().removeRelationship(previewRelationId);
    }
    located.cursor().removeXml();
    if (previewPart != null) {
      ExcelDrawingBinarySupport.cleanupWorkbookImagePartIfUnused(
          sheet.getWorkbook(), previewPart.getPartName());
    }
    return true;
  }

  boolean hasNamedSignatureLine(XSSFSheet sheet, String objectName) {
    return find(sheet, objectName) != null;
  }

  private ExcelSignatureLineSnapshot snapshot(XSSFVMLDrawing vmlDrawing, CTShape shape, int index) {
    String objectName = resolvedName(shape, index);
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
    String signingInstructions;
    try (XmlCursor cursor = signatureLine.newCursor()) {
      signingInstructions =
          ExcelDrawingBinarySupport.nullIfBlank(cursor.getAttributeText(SIGNING_INSTRUCTIONS));
    }
    return new ExcelSignatureLineSnapshot(
        objectName,
        anchor(shape, objectName),
        ExcelDrawingBinarySupport.nullIfBlank(signatureLine.getId()),
        signatureLine.isSetAllowcomments()
            ? STTrueFalse.TRUE.equals(signatureLine.getAllowcomments())
            : null,
        signingInstructions,
        ExcelDrawingBinarySupport.nullIfBlank(signatureLine.getSuggestedsigner()),
        ExcelDrawingBinarySupport.nullIfBlank(signatureLine.getSuggestedsigner2()),
        ExcelDrawingBinarySupport.nullIfBlank(signatureLine.getSuggestedsigneremail()),
        previewFormat,
        previewPart == null ? null : previewPart.getContentType(),
        previewBytes == null ? null : (long) previewBytes.length,
        previewBytes == null ? null : ExcelDrawingBinarySupport.sha256(previewBytes),
        previewDimensions.widthPixels(),
        previewDimensions.heightPixels());
  }

  private LocatedSignatureLine find(XSSFSheet sheet, String objectName) {
    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return java.util.Optional.<LocatedSignatureLine>empty().orElse(null);
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
        if (!resolvedName(shape, index).equals(objectName)) {
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
        ? null
        : new LocatedSignatureLine(matchedShape, matchedCursor, matchedIndex);
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
    if (definition.plainSignature() == null) {
      return;
    }
    signatureLine.setPlainSignature(definition.plainSignature().bytes());
    signatureLine.setContentType(definition.plainSignatureFormat().defaultContentType());
  }

  private static void applyNameAndCommentMetadata(
      CTShape shape, ExcelSignatureLineDefinition definition) {
    shape.setAlt(definition.name());
    applyImageTitle(shape, definition.name());
    shape
        .getSignaturelineArray(0)
        .setAllowcomments(definition.allowComments() ? STTrueFalse.T : STTrueFalse.F);
  }

  static void applyImageTitle(CTShape shape, String name) {
    if (shape.sizeOfImagedataArray() == 0) {
      return;
    }
    shape.getImagedataArray(0).setTitle(name);
  }

  private static void applyIfPresent(String value, java.util.function.Consumer<String> consumer) {
    if (value != null) {
      consumer.accept(value);
    }
  }

  private static List<CTShape> signatureShapes(XSSFVMLDrawing vmlDrawing) {
    List<CTShape> shapes = new ArrayList<>();
    try (XmlCursor cursor = vmlDrawing.getDocument().getXml().newCursor()) {
      for (boolean found = cursor.toFirstChild(); found; found = cursor.toNextSibling()) {
        XmlObject object = cursor.getObject();
        if (object instanceof CTShape shape && shape.sizeOfSignaturelineArray() > 0) {
          shapes.add(shape);
        }
      }
    }
    return List.copyOf(shapes);
  }

  static boolean usesImagePart(
      XSSFSheet sheet, org.apache.poi.openxml4j.opc.PackagePartName imagePartName) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(imagePartName, "imagePartName must not be null");
    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return false;
    }
    for (CTShape shape : signatureShapes(vmlDrawing)) {
      PackagePart previewPart = previewPart(vmlDrawing, shape);
      if (previewPart != null && previewPart.getPartName().equals(imagePartName)) {
        return true;
      }
    }
    return false;
  }

  private static PackagePart previewPart(XSSFVMLDrawing vmlDrawing, CTShape shape) {
    String previewRelationId = previewRelationId(shape);
    return previewRelationId == null
        ? null
        : ExcelDrawingBinarySupport.relatedInternalPart(
            vmlDrawing.getPackagePart(), previewRelationId);
  }

  private static String previewRelationId(CTShape shape) {
    return shape.sizeOfImagedataArray() == 0
        ? null
        : ExcelDrawingBinarySupport.nullIfBlank(shape.getImagedataArray(0).getRelid());
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
      XSSFSheet sheet, CTShape shape, String objectName) {
    if (shape.sizeOfClientDataArray() == 0) {
      throw new IllegalStateException(missingClientDataMessage(sheet, objectName));
    }
    CTClientData clientData = shape.getClientDataArray(0);
    if (clientData.sizeOfAnchorArray() == 0) {
      throw new IllegalStateException(missingAnchorMessage(sheet, objectName));
    }
    return clientData;
  }

  private static String missingClientDataMessage(XSSFSheet sheet, String objectName) {
    return sheet == null
        ? "Signature line '" + objectName + "' is missing VML clientData"
        : "Signature line '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is missing VML clientData";
  }

  private static String missingAnchorMessage(XSSFSheet sheet, String objectName) {
    return sheet == null
        ? "Signature line '" + objectName + "' is missing its VML anchor"
        : "Signature line '"
            + objectName
            + "' on sheet '"
            + sheet.getSheetName()
            + "' is missing its VML anchor";
  }

  private static String resolvedName(CTShape shape, int index) {
    if (shape.sizeOfImagedataArray() > 0) {
      CTImageData imageData = shape.getImagedataArray(0);
      String titledName = ExcelDrawingBinarySupport.nullIfBlank(imageData.getTitle());
      if (titledName != null) {
        return titledName;
      }
    }
    String alt = ExcelDrawingBinarySupport.nullIfBlank(shape.getAlt());
    if (alt != null && !DEFAULT_ALT_TEXT.equals(alt)) {
      return alt;
    }
    CTSignatureLine signatureLine = shape.getSignaturelineArray(0);
    String setupId = ExcelDrawingBinarySupport.nullIfBlank(signatureLine.getId());
    return setupId == null ? "SignatureLine-" + index : "SignatureLine-" + setupId;
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
}
