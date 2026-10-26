package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.chart.ChartInput;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorInput;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectInput;
import dev.erst.gridgrind.core.model.drawing.PictureInput;
import dev.erst.gridgrind.core.model.drawing.ShapeInput;
import dev.erst.gridgrind.core.model.drawing.SignatureLineInput;
import java.util.Objects;

/** Mutation family for drawing-backed workbook objects. */
public sealed interface DrawingMutationAction extends MutationAction {
  /** Creates or replaces one picture-backed drawing object on one sheet. */
  @Variant(
      id = "SET_PICTURE",
      summary = "Create or replace one picture-backed drawing object on one sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.DRAWING_MUTATION)
  record SetPicture(PictureInput picture) implements DrawingMutationAction {
    public SetPicture {
      Objects.requireNonNull(picture, "picture must not be null");
    }
  }

  /** Creates or replaces one signature-line drawing object on one sheet. */
  @Variant(
      id = "SET_SIGNATURE_LINE",
      summary = "Create or replace one signature-line drawing object on one sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.DRAWING_MUTATION)
  record SetSignatureLine(SignatureLineInput signatureLine) implements DrawingMutationAction {
    public SetSignatureLine {
      Objects.requireNonNull(signatureLine, "signatureLine must not be null");
    }
  }

  /** Creates or mutates one supported simple chart on one sheet. */
  @Variant(
      id = "SET_CHART",
      summary = "Create or mutate one supported simple chart on one sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.DRAWING_MUTATION)
  record SetChart(ChartInput chart) implements DrawingMutationAction {
    public SetChart {
      Objects.requireNonNull(chart, "chart must not be null");
    }
  }

  /** Creates or replaces one simple-shape or connector drawing object on one sheet. */
  @Variant(
      id = "SET_SHAPE",
      summary = "Create or replace one simple-shape or connector drawing object on one sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.DRAWING_MUTATION)
  record SetShape(ShapeInput shape) implements DrawingMutationAction {
    public SetShape {
      Objects.requireNonNull(shape, "shape must not be null");
    }
  }

  /** Creates or replaces one embedded-object drawing object on one sheet. */
  @Variant(
      id = "SET_EMBEDDED_OBJECT",
      summary = "Create or replace one embedded-object drawing object on one sheet.",
      targetSelectors = {SheetSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.DRAWING_MUTATION)
  record SetEmbeddedObject(EmbeddedObjectInput embeddedObject) implements DrawingMutationAction {
    public SetEmbeddedObject {
      Objects.requireNonNull(embeddedObject, "embeddedObject must not be null");
    }
  }

  /** Moves one existing drawing object by replacing its anchor authoritatively. */
  @Variant(
      id = "SET_DRAWING_OBJECT_ANCHOR",
      summary = "Move one existing drawing object by replacing its anchor authoritatively.",
      targetSelectors = {DrawingObjectSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.DRAWING_MUTATION)
  record SetDrawingObjectAnchor(DrawingAnchorInput anchor) implements DrawingMutationAction {
    public SetDrawingObjectAnchor {
      Objects.requireNonNull(anchor, "anchor must not be null");
    }
  }

  /** Deletes one existing drawing object by sheet-local name. */
  @Variant(
      id = "DELETE_DRAWING_OBJECT",
      summary = "Delete one existing drawing object by sheet-local name.",
      targetSelectors = {DrawingObjectSelector.ByName.class},
      profile = dev.erst.gridgrind.core.semantics.OperationProfile.DRAWING_MUTATION)
  record DeleteDrawingObject() implements DrawingMutationAction {
    public DeleteDrawingObject {}
  }
}
