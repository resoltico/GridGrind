package dev.erst.gridgrind.excel.drawing;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.microsoft.schemas.office.office.CTSignatureLine;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Optional;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STTrueFalse;

/** Focused branch coverage for signature-line setup and preview extraction helpers. */
class ExcelSignatureLineSnapshotSupportTest {
  @Test
  void signatureLineSetupReturnsEmptyOnlyWhenAllMetadataIsAbsent() {
    assertTrue(
        ExcelSignatureLineSnapshotSupport.signatureLineSetup(
                CTSignatureLine.Factory.newInstance(), null)
            .isEmpty());

    CTSignatureLine allowComments = CTSignatureLine.Factory.newInstance();
    allowComments.setAllowcomments(STTrueFalse.TRUE);
    assertEquals(
        Optional.of(true),
        ExcelSignatureLineSnapshotSupport.signatureLineSetup(allowComments, null)
            .orElseThrow()
            .allowComments());

    CTSignatureLine setupId = CTSignatureLine.Factory.newInstance();
    setupId.setId("Setup-42");
    assertEquals(
        Optional.of("Setup-42"),
        ExcelSignatureLineSnapshotSupport.signatureLineSetup(setupId, null)
            .orElseThrow()
            .setupId());

    CTSignatureLine suggestedSigner = CTSignatureLine.Factory.newInstance();
    suggestedSigner.setSuggestedsigner("Ada Lovelace");
    assertEquals(
        Optional.of("Ada Lovelace"),
        ExcelSignatureLineSnapshotSupport.signatureLineSetup(suggestedSigner, null)
            .orElseThrow()
            .suggestedSigner());

    CTSignatureLine suggestedSigner2 = CTSignatureLine.Factory.newInstance();
    suggestedSigner2.setSuggestedsigner2("Finance");
    assertEquals(
        Optional.of("Finance"),
        ExcelSignatureLineSnapshotSupport.signatureLineSetup(suggestedSigner2, null)
            .orElseThrow()
            .suggestedSigner2());

    CTSignatureLine suggestedSignerEmail = CTSignatureLine.Factory.newInstance();
    suggestedSignerEmail.setSuggestedsigneremail("ada@example.com");
    assertEquals(
        Optional.of("ada@example.com"),
        ExcelSignatureLineSnapshotSupport.signatureLineSetup(suggestedSignerEmail, null)
            .orElseThrow()
            .suggestedSignerEmail());

    assertEquals(
        Optional.of("Review before signing."),
        ExcelSignatureLineSnapshotSupport.signatureLineSetup(
                CTSignatureLine.Factory.newInstance(), "Review before signing.")
            .orElseThrow()
            .signingInstructions());
  }

  @Test
  void signatureLinePreviewRequiresFormatPartAndBytes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      PackagePart previewPart =
          workbook
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/media/signature-preview.png"),
                  "image/png");
      ExcelDrawingSnapshotSupport.RasterDimensions dimensions =
          new ExcelDrawingSnapshotSupport.RasterDimensions(320, 180);
      byte[] previewBytes = new byte[] {1, 2, 3, 4};

      assertTrue(
          ExcelSignatureLineSnapshotSupport.signatureLinePreview(
                  null, previewPart, previewBytes, dimensions)
              .isEmpty());
      assertTrue(
          ExcelSignatureLineSnapshotSupport.signatureLinePreview(
                  ExcelPictureFormat.PNG, null, previewBytes, dimensions)
              .isEmpty());
      assertTrue(
          ExcelSignatureLineSnapshotSupport.signatureLinePreview(
                  ExcelPictureFormat.PNG, previewPart, null, dimensions)
              .isEmpty());

      ExcelSignatureLineSnapshot.Preview preview =
          ExcelSignatureLineSnapshotSupport.signatureLinePreview(
                  ExcelPictureFormat.PNG, previewPart, previewBytes, dimensions)
              .orElseThrow();
      assertEquals("image/png", preview.contentType());
      assertEquals(4L, preview.byteSize());
      assertEquals(Optional.of(320), preview.widthPixels());
      assertEquals(Optional.of(180), preview.heightPixels());
      assertEquals(Optional.of(ExcelDrawingBinarySupport.sha256(previewBytes)), preview.sha256());
    }
  }
}
