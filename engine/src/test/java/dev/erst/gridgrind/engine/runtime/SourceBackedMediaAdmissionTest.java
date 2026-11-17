package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.model.drawing.*;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.semantics.ExecutionConventions;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.media.MediaFixtures;
import dev.erst.gridgrind.engine.poi.NativeInvocationContext;
import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * Source/native handoff retains occurrence origins, one acquisition and per-occurrence media work.
 */
class SourceBackedMediaAdmissionTest {
  @TempDir Path root;
  private static final DrawingAnchorSpec.TwoCell ANCHOR =
      DrawingAnchorSpec.TwoCell.moveAndResize(
          new DrawingMarkerSpec(0, 0), new DrawingMarkerSpec(2, 2));

  @ParameterizedTest
  @CsvSource({
    "PICTURE,INLINE,steps[0].action.picture.image.source.base64Data",
    "PICTURE,FILE,steps[0].action.picture.image.source.path",
    "PICTURE,STDIN,steps[0].action.picture.image.source.type",
    "EMBEDDED,INLINE,steps[0].action.embeddedObject.previewImage.source.base64Data",
    "EMBEDDED,FILE,steps[0].action.embeddedObject.previewImage.source.path",
    "EMBEDDED,STDIN,steps[0].action.embeddedObject.previewImage.source.type",
    "SIGNATURE,INLINE,steps[0].action.signatureLine.plainSignature.source.base64Data",
    "SIGNATURE,FILE,steps[0].action.signatureLine.plainSignature.source.path",
    "SIGNATURE,STDIN,steps[0].action.signatureLine.plainSignature.source.type"
  })
  void everyAuthoredRouteKeepsItsOriginalSourceLocation(
      String route, String kind, String expectedPath) throws Exception {
    byte[] malformed = {1, 2, 3};
    Path file = root.resolve("preview.bin");
    Files.write(file, malformed);
    BinarySourceSpec source =
        switch (kind) {
          case "INLINE" -> new BinarySourceSpec.InlineBase64("AQID");
          case "FILE" -> new BinarySourceSpec.File("preview.bin");
          default -> new BinarySourceSpec.Stdin();
        };
    var step =
        new MutationStep(
            new StepId("image"),
            new SheetSelector.ByName(new SheetName("Sheet1")),
            action(route, source));
    var request = plan(List.of(step));
    var inputs =
        bindings(
            List.of(
                new ExecutionGrant.ReadAuthority.File(file),
                new ExecutionGrant.ReadAuthority.Stdin()),
            malformed,
            new AtomicInteger());
    try (var context = NativeInvocationContext.enterOwned(ExecutionConventions.STANDARD)) {
      context.requireOpen();
      var result = RequestPreflight.verify(request, inputs);
      assertTrue(result.preparedRequest().isEmpty());
      assertEquals(1, result.problems().size(), result.problems().toString());
      var problem = result.problems().getFirst();
      assertEquals(GridGrindProblemCode.INVALID_REQUEST, problem.code());
      var location = assertInstanceOf(ProblemContext.ResolveInputs.class, problem.context());
      assertEquals(expectedPath, location.json().orElseThrow().jsonPathValue().orElseThrow());
      if ("FILE".equals(kind)) assertEquals(Optional.of("preview.bin"), location.inputPath());
      assertSame(step, request.steps().getFirst());
    } finally {
      inputs.closeRunResources();
    }
  }

  @Test
  void maximumStepPlanCountsEveryUseOfTheSameStdinSnapshot() throws Exception {
    byte[] png = MediaFixtures.validPng();
    assertEquals(85, png.length);
    var opens = new AtomicInteger();
    var inputs = bindings(List.of(new ExecutionGrant.ReadAuthority.Stdin()), png, opens);
    BinarySourceSpec shared = new BinarySourceSpec.Stdin();
    var target = new SheetSelector.ByName(new SheetName("Sheet1"));
    var image = action("PICTURE", shared);
    List<WorkbookStep> steps = new ArrayList<>();
    steps.add(
        new MutationStep(new StepId("sheet"), target, new WorkbookMutationAction.EnsureSheet()));
    for (int i = 1; i < 10_000; i++)
      steps.add(new MutationStep(new StepId("image-" + i), target, image));
    try (var context = NativeInvocationContext.enterOwned(ExecutionConventions.STANDARD)) {
      context.requireOpen();
      var result = RequestPreflight.verify(plan(steps), inputs);
      assertTrue(result.problems().isEmpty(), result.problems().toString());
      assertEquals(10_000, result.preparedRequest().orElseThrow().plan().steps().size());
      assertEquals(849_915, inputs.mediaBudget().mediaUseBytes());
      assertEquals(1, opens.get());
    } finally {
      inputs.closeRunResources();
    }
  }

  private ExecutionInputBindings bindings(
      List<ExecutionGrant.ReadAuthority> reads, byte[] stdin, AtomicInteger opens) {
    var grant =
        new ExecutionGrant(
            4096,
            reads,
            List.of("SET_PICTURE", "SET_EMBEDDED_OBJECT", "SET_SIGNATURE_LINE", "ENSURE_SHEET"),
            new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
            new ExecutionGrant.PublicationAuthority.None(),
            List.of(),
            AcceptancePolicy.minimum());
    return new ExecutionInputBindings(
        root,
        root.resolve("scratch"),
        grant,
        Optional.of(() -> {
          opens.incrementAndGet();
          return new ByteArrayInputStream(stdin);
        }),
        Optional.empty());
  }

  private static DrawingMutationAction action(String route, BinarySourceSpec source) {
    var image = new PictureDataSpec(ExcelPictureFormat.PNG, source);
    return switch (route) {
      case "PICTURE" ->
          new DrawingMutationAction.SetPicture(
              new PictureSpec("Image", image, ANCHOR, Optional.empty()));
      case "EMBEDDED" ->
          new DrawingMutationAction.SetEmbeddedObject(
              new EmbeddedObjectSpec(
                  "Object",
                  "Payload",
                  "file.bin",
                  "open",
                  new BinarySourceSpec.InlineBase64("BA=="),
                  image,
                  ANCHOR));
      default ->
          new DrawingMutationAction.SetSignatureLine(
              new SignatureLineSpec(
                  "Signature",
                  ANCHOR,
                  false,
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.of("Synthetic signer"),
                  Optional.empty(),
                  Optional.of(image)));
    };
  }

  private static WorkbookPlan plan(List<WorkbookStep> steps) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicySpec.defaults(),
        FormulaEnvironmentSpec.empty(),
        steps);
  }
}
