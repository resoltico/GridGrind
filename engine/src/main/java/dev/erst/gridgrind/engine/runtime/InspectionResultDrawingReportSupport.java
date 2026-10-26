package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectPayload;
import java.util.Base64;

/** Converts drawing and chart workbook snapshots into protocol report records. */
final class InspectionResultDrawingReportSupport {
  private InspectionResultDrawingReportSupport() {}

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
