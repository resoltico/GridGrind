package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
import java.util.Base64;

/** Converts drawing and chart workbook snapshots into protocol report records. */
final class InspectionResultDrawingReportSupport {
  private InspectionResultDrawingReportSupport() {}

  static DrawingObjectReport toDrawingObjectReport(ExcelDrawingObjectSnapshot snapshot) {
    return switch (snapshot) {
      case ExcelDrawingObjectSnapshot.Picture picture ->
          new DrawingObjectReport.Picture(
              picture.name(),
              picture.anchor(),
              picture.format(),
              picture.contentType(),
              picture.byteSize(),
              picture.sha256(),
              picture.widthPixels(),
              picture.heightPixels(),
              picture.description());
      case ExcelDrawingObjectSnapshot.Chart chart ->
          new DrawingObjectReport.Chart(
              chart.name(),
              chart.anchor(),
              chart.supported(),
              chart.plotTypeTokens(),
              chart.title());
      case ExcelDrawingObjectSnapshot.Shape shape ->
          new DrawingObjectReport.Shape(
              shape.name(),
              shape.anchor(),
              shape.kind(),
              shape.presetGeometryToken(),
              shape.text(),
              shape.childCount());
      case ExcelDrawingObjectSnapshot.EmbeddedObject embeddedObject ->
          new DrawingObjectReport.EmbeddedObject(
              embeddedObject.name(),
              embeddedObject.anchor(),
              embeddedObject.packagingKind(),
              embeddedObject.label(),
              embeddedObject.fileName(),
              embeddedObject.command(),
              embeddedObject.contentType(),
              embeddedObject.byteSize(),
              embeddedObject.sha256(),
              embeddedObject.previewFormat(),
              embeddedObject.previewByteSize(),
              embeddedObject.previewSha256());
      case ExcelDrawingObjectSnapshot.SignatureLine signatureLine ->
          new DrawingObjectReport.SignatureLine(
              signatureLine.name(),
              signatureLine.anchor(),
              signatureLine
                  .setup()
                  .map(
                      setup ->
                          new DrawingObjectReport.SignatureSetup(
                              setup.setupId(),
                              setup.allowComments(),
                              setup.signingInstructions(),
                              setup.suggestedSigner(),
                              setup.suggestedSigner2(),
                              setup.suggestedSignerEmail())),
              signatureLine
                  .preview()
                  .map(
                      preview ->
                          new DrawingObjectReport.SignaturePreview(
                              preview.format(),
                              preview.contentType(),
                              preview.byteSize(),
                              preview.sha256(),
                              preview.widthPixels(),
                              preview.heightPixels())));
    };
  }

  static DrawingObjectPayloadReport toDrawingObjectPayloadReport(
      ExcelDrawingObjectPayload payload) {
    return switch (payload) {
      case ExcelDrawingObjectPayload.Picture picture ->
          new DrawingObjectPayloadReport.Picture(
              picture.name(),
              picture.format(),
              picture.contentType(),
              picture.fileName(),
              picture.sha256(),
              Base64.getEncoder().encodeToString(picture.data().bytes()),
              picture.description());
      case ExcelDrawingObjectPayload.EmbeddedObject embeddedObject ->
          new DrawingObjectPayloadReport.EmbeddedObject(
              embeddedObject.name(),
              embeddedObject.packagingKind(),
              embeddedObject.contentType(),
              embeddedObject.fileName(),
              embeddedObject.sha256(),
              Base64.getEncoder().encodeToString(embeddedObject.data().bytes()),
              embeddedObject.label(),
              embeddedObject.command());
    };
  }
}
