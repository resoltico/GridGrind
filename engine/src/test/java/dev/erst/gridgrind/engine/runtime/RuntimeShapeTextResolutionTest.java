package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;

import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies standard-input materialization for drawing-shape text. */
class RuntimeShapeTextResolutionTest {
  @Test
  void sourceBackedStructuredResolutionBindsShapeTextFromStandardInput() throws Exception {
    ShapeInput.SimpleShape shape =
        new ShapeInput.SimpleShape(
            "Queue Banner", anchor(), "rect", Optional.of(TextSourceInput.standardInput()));
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            Path.of("tmp", "runtime-residual-shape"),
            Path.of("tmp", "runtime-residual-shape", "temp-root"),
            "Queue ready".getBytes(StandardCharsets.UTF_8),
            ExecutionGrantTestSupport.noPublicationWithStandardInput(List.of()));

    ShapeInput resolved = SourceBackedStructuredInputResolver.resolveShape(shape, bindings);
    ShapeInput connector = new ShapeInput.Connector("Connector", anchor());
    ShapeInput unresolvedConnector =
        SourceBackedStructuredInputResolver.resolveShape(connector, bindings);
    ShapeInput unchangedShape =
        new ShapeInput.SimpleShape(
            "Inline Banner", anchor(), "rect", Optional.of(TextSourceInput.inline("Ready")));
    ShapeInput unresolvedShape =
        SourceBackedStructuredInputResolver.resolveShape(unchangedShape, bindings);
    ShapeInput noTextShape =
        new ShapeInput.SimpleShape("Silent Banner", anchor(), "rect", Optional.empty());
    ShapeInput unresolvedNoTextShape =
        SourceBackedStructuredInputResolver.resolveShape(noTextShape, bindings);

    assertInstanceOf(ShapeInput.SimpleShape.class, resolved);
    assertNotSame(shape, resolved);
    assertEquals(
        Optional.of(TextSourceInput.inline("Queue ready")),
        ((ShapeInput.SimpleShape) resolved).text());
    assertEquals(connector, unresolvedConnector);
    assertSame(unchangedShape, unresolvedShape);
    assertSame(noTextShape, unresolvedNoTextShape);
  }

  private static DrawingAnchorInput.TwoCell anchor() {
    return new DrawingAnchorInput.TwoCell(
        new DrawingMarkerInput(0, 0),
        new DrawingMarkerInput(2, 2),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }
}
