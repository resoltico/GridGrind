package dev.erst.gridgrind.excel.drawing;

import com.microsoft.schemas.office.office.CTSignatureLine;
import com.microsoft.schemas.vml.CTImageData;
import com.microsoft.schemas.vml.CTShape;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import javax.xml.namespace.QName;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STTrueFalse;

/** Shared snapshot assembly support for signature-line reads. */
final class ExcelSignatureLineSnapshotSupport {
  private static final String DEFAULT_ALT_TEXT = "Microsoft Office Signature Line...";
  private static final String MS_OFFICE_URN = "urn:schemas-microsoft-com:office:office";
  private static final QName SIGNING_INSTRUCTIONS = new QName(MS_OFFICE_URN, "signinginstructions");

  private ExcelSignatureLineSnapshotSupport() {}

  static List<CTShape> signatureShapes(XSSFVMLDrawing vmlDrawing) {
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

  static String resolvedName(CTShape shape, int index) {
    if (shape.sizeOfImagedataArray() > 0) {
      CTImageData imageData = shape.getImagedataArray(0);
      String titledName =
          ExcelDrawingBinarySupport.blankAsOptional(imageData.getTitle()).orElse(null);
      if (titledName != null) {
        return titledName;
      }
    }
    String alt = ExcelDrawingBinarySupport.blankAsOptional(shape.getAlt()).orElse(null);
    if (alt != null && !DEFAULT_ALT_TEXT.equals(alt)) {
      return alt;
    }
    CTSignatureLine signatureLine = shape.getSignaturelineArray(0);
    String setupId = ExcelDrawingBinarySupport.blankAsOptional(signatureLine.getId()).orElse(null);
    return setupId == null ? "SignatureLine-" + index : "SignatureLine-" + setupId;
  }

  static @Nullable String signingInstructions(CTSignatureLine signatureLine) {
    try (XmlCursor cursor = signatureLine.newCursor()) {
      return ExcelDrawingBinarySupport.blankAsOptional(
              cursor.getAttributeText(SIGNING_INSTRUCTIONS))
          .orElse(null);
    }
  }

  static Optional<ExcelSignatureLineSnapshot.Setup> signatureLineSetup(
      CTSignatureLine signatureLine, @Nullable String signingInstructions) {
    Optional<String> setupId = ExcelDrawingBinarySupport.blankAsOptional(signatureLine.getId());
    Optional<Boolean> allowComments =
        signatureLine.isSetAllowcomments()
            ? Optional.of(STTrueFalse.TRUE.equals(signatureLine.getAllowcomments()))
            : Optional.empty();
    Optional<String> suggestedSigner =
        ExcelDrawingBinarySupport.blankAsOptional(signatureLine.getSuggestedsigner());
    Optional<String> suggestedSigner2 =
        ExcelDrawingBinarySupport.blankAsOptional(signatureLine.getSuggestedsigner2());
    Optional<String> suggestedSignerEmail =
        ExcelDrawingBinarySupport.blankAsOptional(signatureLine.getSuggestedsigneremail());
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
        new ExcelSignatureLineSnapshot.Setup(
            setupId,
            allowComments,
            signingInstructionsValue,
            suggestedSigner,
            suggestedSigner2,
            suggestedSignerEmail));
  }

  static Optional<ExcelSignatureLineSnapshot.Preview> signatureLinePreview(
      @Nullable ExcelPictureFormat previewFormat,
      @Nullable PackagePart previewPart,
      byte @Nullable [] previewBytes,
      ExcelDrawingSnapshotSupport.RasterDimensions previewDimensions) {
    if (previewFormat == null || previewPart == null || previewBytes == null) {
      return Optional.empty();
    }
    return Optional.of(
        new ExcelSignatureLineSnapshot.Preview(
            previewFormat,
            previewPart.getContentType(),
            previewBytes.length,
            Optional.of(ExcelDrawingBinarySupport.sha256(previewBytes)),
            Optional.ofNullable(previewDimensions.widthPixels()),
            Optional.ofNullable(previewDimensions.heightPixels())));
  }
}
