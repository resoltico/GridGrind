package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Drawing-object commands for pictures, charts, shapes, and embedded payloads. */
public sealed interface WorkbookDrawingCommand extends WorkbookCommand
    permits WorkbookDrawingCommand.SetPicture,
        WorkbookDrawingCommand.SetSignatureLine,
        WorkbookDrawingCommand.SetChart,
        WorkbookDrawingCommand.SetShape,
        WorkbookDrawingCommand.SetEmbeddedObject,
        WorkbookDrawingCommand.SetDrawingObjectAnchor,
        WorkbookDrawingCommand.DeleteDrawingObject {

  /** Creates or replaces one picture-backed drawing object on a single sheet. */
  record SetPicture(String sheetName, ExcelPictureDefinition picture)
      implements WorkbookDrawingCommand {
    public SetPicture {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(picture, "picture must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or replaces one signature-line drawing object on a single sheet. */
  record SetSignatureLine(String sheetName, ExcelSignatureLineDefinition signatureLine)
      implements WorkbookDrawingCommand {
    public SetSignatureLine {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(signatureLine, "signatureLine must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or mutates one supported simple chart on a single sheet. */
  record SetChart(String sheetName, ExcelChartDefinition chart) implements WorkbookDrawingCommand {
    public SetChart {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(chart, "chart must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or replaces one simple-shape or connector drawing object on a single sheet. */
  record SetShape(String sheetName, ExcelShapeDefinition shape) implements WorkbookDrawingCommand {
    public SetShape {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(shape, "shape must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Creates or replaces one embedded-object drawing object on a single sheet. */
  record SetEmbeddedObject(String sheetName, ExcelEmbeddedObjectDefinition embeddedObject)
      implements WorkbookDrawingCommand {
    public SetEmbeddedObject {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(embeddedObject, "embeddedObject must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
    }
  }

  /** Moves one existing drawing object by replacing its anchor authoritatively. */
  record SetDrawingObjectAnchor(
      String sheetName, String objectName, ExcelDrawingAnchor.TwoCell anchor)
      implements WorkbookDrawingCommand {
    public SetDrawingObjectAnchor {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(objectName, "objectName must not be null");
      Objects.requireNonNull(anchor, "anchor must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (objectName.isBlank()) {
        throw new IllegalArgumentException("objectName must not be blank");
      }
    }
  }

  /** Deletes one existing drawing object by sheet-local name. */
  record DeleteDrawingObject(String sheetName, String objectName)
      implements WorkbookDrawingCommand {
    public DeleteDrawingObject {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(objectName, "objectName must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (objectName.isBlank()) {
        throw new IllegalArgumentException("objectName must not be blank");
      }
    }
  }
}
