package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies standard-input materialization for drawing-shape text. */
class RuntimeShapeTextResolutionTest {
  @Test
  void sourceBackedStructuredResolutionBindsShapeTextFromStandardInput() throws Exception {
    ShapeSpec.SimpleShape shape =
        new ShapeSpec.SimpleShape(
            "Queue Banner", anchor(), "rect", Optional.of(TextSourceSpec.standardInput()));
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Path.of("tmp", "runtime-residual-shape"),
            Path.of("tmp", "runtime-residual-shape", "temp-root"),
            "Queue ready".getBytes(StandardCharsets.UTF_8),
            ExecutionGrantTestSupport.noPublicationWithStandardInput(List.of()));

    ShapeSpec resolved = SourceBackedStructuredInputResolver.resolveShape(shape, bindings);
    ShapeSpec connector = new ShapeSpec.Connector("Connector", anchor());
    ShapeSpec unresolvedConnector =
        SourceBackedStructuredInputResolver.resolveShape(connector, bindings);
    ShapeSpec unchangedShape =
        new ShapeSpec.SimpleShape(
            "Inline Banner", anchor(), "rect", Optional.of(TextSourceSpec.inline("Ready")));
    ShapeSpec unresolvedShape =
        SourceBackedStructuredInputResolver.resolveShape(unchangedShape, bindings);
    ShapeSpec noTextShape =
        new ShapeSpec.SimpleShape("Silent Banner", anchor(), "rect", Optional.empty());
    ShapeSpec unresolvedNoTextShape =
        SourceBackedStructuredInputResolver.resolveShape(noTextShape, bindings);

    assertInstanceOf(ShapeSpec.SimpleShape.class, resolved);
    assertNotSame(shape, resolved);
    assertEquals(
        Optional.of(TextSourceSpec.inline("Queue ready")),
        ((ShapeSpec.SimpleShape) resolved).text());
    assertEquals(connector, unresolvedConnector);
    assertSame(unchangedShape, unresolvedShape);
    assertSame(noTextShape, unresolvedNoTextShape);
  }

  private static DrawingAnchorSpec.TwoCell anchor() {
    return new DrawingAnchorSpec.TwoCell(
        new DrawingMarkerSpec(0, 0),
        new DrawingMarkerSpec(2, 2),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }
}
