package dev.erst.gridgrind.engine.drawing;

import com.microsoft.schemas.office.excel.CTClientData;
import com.microsoft.schemas.office.office.CTSignatureLine;
import com.microsoft.schemas.vml.CTImageData;
import com.microsoft.schemas.vml.CTShape;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.drawing.ResolvedSignatureLine;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;
import javax.xml.namespace.QName;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSignatureLine;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STTrueFalse;

/** Owns VML signature-line mutation, metadata, previews and anchors. */
@SuppressWarnings("PMD.CommentRequired")
public final class SignatureAdapter {

  private static final Pattern VML_ANCHOR_SEPARATOR = Pattern.compile("\\s*,\\s*");

  public List<DrawingObjectReport.SignatureLine> signatureLines(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    XSSFVMLDrawing vmlDrawing = sheet.getVMLDrawing(false);
    if (vmlDrawing == null) {
      return List.of();
    }

    List<DrawingObjectReport.SignatureLine> snapshots = new ArrayList<>();
    int index = 1;
    for (CTShape shape : signatureShapes(vmlDrawing)) {
      int currentIndex = index;
      index++;
      snapshots.add(snapshot(vmlDrawing, shape, currentIndex));
    }
    return List.copyOf(snapshots);
  }

  public void setSignatureLine(XSSFSheet sheet, ResolvedSignatureLine definition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    deleteIfPresent(sheet, definition.name());

    XSSFSignatureLine signatureLine = configuredSignatureLine(definition);
    signatureLine.add(sheet, toPoiAnchor(definition.anchor()));
    applyNameAndCommentMetadata(signatureLine.getSignatureShape(), definition);
  }

  public boolean updateAnchorIfPresent(
      XSSFSheet sheet, String objectName, DrawingAnchorSpec.TwoCell anchor) {
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
      DrawingAdapter.cleanupWorkbookImagePartIfUnused(
          sheet.getWorkbook(), previewPart.getPartName());
    }
    return true;
  }

  public boolean hasNamedSignatureLine(XSSFSheet sheet, String objectName) {
    return find(sheet, objectName).isPresent();
  }

  private DrawingObjectReport.SignatureLine snapshot(
      XSSFVMLDrawing vmlDrawing, CTShape shape, int index) {
    String objectName = resolvedName(shape, index);
    PackagePart previewPart = previewPart(vmlDrawing, shape);
    ExcelPictureFormat previewFormat =
        previewPart == null
            ? null
            : ExcelPictureFormat.fromContentType(previewPart.getContentType());
    byte[] previewBytes = previewPart == null ? null : DrawingAdapter.partBytes(previewPart);
    DrawingAdapter.RasterDimensions previewDimensions =
        previewBytes == null
            ? DrawingAdapter.RasterDimensions.none()
            : DrawingAdapter.rasterDimensions(previewBytes);
    CTSignatureLine signatureLine = shape.getSignaturelineArray(0);
    return new DrawingObjectReport.SignatureLine(
        objectName,
        anchor(shape, objectName),
        signatureLineSetup(signatureLine, signingInstructions(signatureLine)),
        signatureLinePreview(previewFormat, previewPart, previewBytes, previewDimensions));
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
        ? Optional.empty()
        : Optional.of(
            new LocatedSignatureLine(
                matchedShape,
                Objects.requireNonNull(matchedCursor, "matchedCursor must not be null"),
                matchedIndex));
  }

  private static XSSFSignatureLine configuredSignatureLine(ResolvedSignatureLine definition) {
    XSSFSignatureLine signatureLine = new XSSFSignatureLine();
    applyOptionalMetadata(signatureLine, definition);
    applyPlainSignature(signatureLine, definition);
    return signatureLine;
  }

  private static void applyOptionalMetadata(
      XSSFSignatureLine signatureLine, ResolvedSignatureLine definition) {
    definition.signingInstructions().ifPresent(signatureLine::setSigningInstructions);
    definition.suggestedSigner().ifPresent(signatureLine::setSuggestedSigner);
    definition.suggestedSigner2().ifPresent(signatureLine::setSuggestedSigner2);
    definition.suggestedSignerEmail().ifPresent(signatureLine::setSuggestedSignerEmail);
    definition.caption().ifPresent(signatureLine::setCaption);
    definition.invalidStamp().ifPresent(signatureLine::setInvalidStamp);
  }

  private static void applyPlainSignature(
      XSSFSignatureLine signatureLine, ResolvedSignatureLine definition) {
    if (definition.plainSignature().isEmpty()) {
      return;
    }
    signatureLine.setPlainSignature(definition.plainSignature().orElseThrow().bytes());
    signatureLine.setContentType(
        definition.plainSignatureFormat().orElseThrow().defaultContentType());
  }

  private static void applyNameAndCommentMetadata(CTShape shape, ResolvedSignatureLine definition) {
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

  public static boolean usesImagePart(
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

  private static @Nullable PackagePart previewPart(XSSFVMLDrawing vmlDrawing, CTShape shape) {
    String previewRelationId = previewRelationId(shape);
    return previewRelationId == null
        ? null
        : DrawingAdapter.relatedInternalPart(vmlDrawing.getPackagePart(), previewRelationId)
            .orElse(null);
  }

  private static @Nullable String previewRelationId(CTShape shape) {
    return shape.sizeOfImagedataArray() == 0
        ? null
        : DrawingAdapter.blankAsOptional(shape.getImagedataArray(0).getRelid()).orElse(null);
  }

  private static DrawingAnchorReport.TwoCell anchor(CTShape shape, String objectName) {
    String value = requiredClientData(null, shape, objectName).getAnchorArray(0);
    String[] tokens = VML_ANCHOR_SEPARATOR.split(value, -1);
    if (tokens.length != 8) {
      throw new IllegalStateException(
          "Signature line '" + objectName + "' is backed by an invalid VML anchor: " + value);
    }
    try {
      return new DrawingAnchorReport.TwoCell(
          new DrawingMarkerReport(
              Integer.parseInt(tokens[0]),
              Integer.parseInt(tokens[2]),
              Integer.parseInt(tokens[1]),
              Integer.parseInt(tokens[3])),
          new DrawingMarkerReport(
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

  private static String anchorString(DrawingAnchorSpec.TwoCell anchor) {
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

  private static XSSFClientAnchor toPoiAnchor(DrawingAnchorSpec.TwoCell anchor) {
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

  private static final String DEFAULT_ALT_TEXT = "Microsoft Office Signature Line...";
  private static final String MS_OFFICE_URN = "urn:schemas-microsoft-com:office:office";
  private static final QName SIGNING_INSTRUCTIONS = new QName(MS_OFFICE_URN, "signinginstructions");

  public static List<CTShape> signatureShapes(XSSFVMLDrawing vmlDrawing) {
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

  public static String resolvedName(CTShape shape, int index) {
    if (shape.sizeOfImagedataArray() > 0) {
      CTImageData imageData = shape.getImagedataArray(0);
      String titledName = DrawingAdapter.blankAsOptional(imageData.getTitle()).orElse(null);
      if (titledName != null) {
        return titledName;
      }
    }
    String alt = DrawingAdapter.blankAsOptional(shape.getAlt()).orElse(null);
    if (alt != null && !DEFAULT_ALT_TEXT.equals(alt)) {
      return alt;
    }
    CTSignatureLine signatureLine = shape.getSignaturelineArray(0);
    String setupId = DrawingAdapter.blankAsOptional(signatureLine.getId()).orElse(null);
    return setupId == null ? "SignatureLine-" + index : "SignatureLine-" + setupId;
  }

  public static @Nullable String signingInstructions(CTSignatureLine signatureLine) {
    try (XmlCursor cursor = signatureLine.newCursor()) {
      return DrawingAdapter.blankAsOptional(cursor.getAttributeText(SIGNING_INSTRUCTIONS))
          .orElse(null);
    }
  }

  public static Optional<DrawingObjectReport.SignatureSetup> signatureLineSetup(
      CTSignatureLine signatureLine, @Nullable String signingInstructions) {
    Optional<String> setupId = DrawingAdapter.blankAsOptional(signatureLine.getId());
    Optional<Boolean> allowComments =
        signatureLine.isSetAllowcomments()
            ? Optional.of(STTrueFalse.TRUE.equals(signatureLine.getAllowcomments()))
            : Optional.empty();
    Optional<String> suggestedSigner =
        DrawingAdapter.blankAsOptional(signatureLine.getSuggestedsigner());
    Optional<String> suggestedSigner2 =
        DrawingAdapter.blankAsOptional(signatureLine.getSuggestedsigner2());
    Optional<String> suggestedSignerEmail =
        DrawingAdapter.blankAsOptional(signatureLine.getSuggestedsigneremail());
    Optional<String> signingInstructionsValue = Optional.ofNullable(signingInstructions);
    if (setupId.isEmpty()
        && allowComments.isEmpty()
        && signingInstructionsValue.isEmpty()
        && suggestedSigner.isEmpty()
        && suggestedSigner2.isEmpty()
        && suggestedSignerEmail.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(
        new DrawingObjectReport.SignatureSetup(
            setupId,
            allowComments,
            signingInstructionsValue,
            suggestedSigner,
            suggestedSigner2,
            suggestedSignerEmail));
  }

  public static Optional<DrawingObjectReport.SignaturePreview> signatureLinePreview(
      @Nullable ExcelPictureFormat previewFormat,
      @Nullable PackagePart previewPart,
      byte @Nullable [] previewBytes,
      DrawingAdapter.RasterDimensions previewDimensions) {
    if (previewFormat == null || previewPart == null || previewBytes == null) {
      return Optional.empty();
    }
    return Optional.of(
        new DrawingObjectReport.SignaturePreview(
            previewFormat,
            previewPart.getContentType(),
            previewBytes.length,
            Optional.of(DrawingAdapter.sha256(previewBytes)),
            Optional.ofNullable(previewDimensions.widthPixels()),
            Optional.ofNullable(previewDimensions.heightPixels())));
  }
}
