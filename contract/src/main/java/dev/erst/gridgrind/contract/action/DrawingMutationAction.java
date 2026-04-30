package dev.erst.gridgrind.contract.action;

import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import java.util.Objects;

/** Mutation family for drawing-backed workbook objects. */
public sealed interface DrawingMutationAction extends MutationAction {
  /** Creates or replaces one picture-backed drawing object on one sheet. */
  record SetPicture(PictureInput picture) implements DrawingMutationAction {
    public SetPicture {
      Objects.requireNonNull(picture, "picture must not be null");
    }
  }

  /** Creates or replaces one signature-line drawing object on one sheet. */
  record SetSignatureLine(SignatureLineInput signatureLine) implements DrawingMutationAction {
    public SetSignatureLine {
      Objects.requireNonNull(signatureLine, "signatureLine must not be null");
    }
  }

  /** Creates or mutates one supported simple chart on one sheet. */
  record SetChart(ChartInput chart) implements DrawingMutationAction {
    public SetChart {
      Objects.requireNonNull(chart, "chart must not be null");
    }
  }

  /** Creates or replaces one simple-shape or connector drawing object on one sheet. */
  record SetShape(ShapeInput shape) implements DrawingMutationAction {
    public SetShape {
      Objects.requireNonNull(shape, "shape must not be null");
    }
  }

  /** Creates or replaces one embedded-object drawing object on one sheet. */
  record SetEmbeddedObject(EmbeddedObjectInput embeddedObject) implements DrawingMutationAction {
    public SetEmbeddedObject {
      Objects.requireNonNull(embeddedObject, "embeddedObject must not be null");
    }
  }

  /** Moves one existing drawing object by replacing its anchor authoritatively. */
  record SetDrawingObjectAnchor(DrawingAnchorInput anchor) implements DrawingMutationAction {
    public SetDrawingObjectAnchor {
      Objects.requireNonNull(anchor, "anchor must not be null");
    }
  }

  /** Deletes one existing drawing object by sheet-local name. */
  record DeleteDrawingObject() implements DrawingMutationAction {
    public DeleteDrawingObject {}
  }
}
