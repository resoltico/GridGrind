package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.PreparedMutationStep;
import dev.erst.gridgrind.core.plan.PreparedStep;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.MediaStructureReason;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.engine.backend.xssf.UnexpectedNativeContentException;
import dev.erst.gridgrind.engine.media.AuthoredImageAdmission;
import dev.erst.gridgrind.engine.media.MediaAdmissionBudget;
import dev.erst.gridgrind.engine.media.MediaIssues;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.Optional;

/**
 * One media occurrence: bounded source capture/reservation, original origin, and prepared
 * inspection.
 */
final class MediaMaterializationUse {
  private final MediaAdmissionBudget budget;
  private long reserved = -1;
  private Optional<BinarySourceSpec> original = Optional.empty();
  private Optional<JsonLocation> location = Optional.empty();

  MediaMaterializationUse(MediaAdmissionBudget budget) {
    this.budget = budget;
  }

  InputSnapshotLimit readLimit() {
    long maximum = budget.payloadAllowance();
    return new InputSnapshotLimit(maximum, MediaIssues.limit(budget.payloadLimitId()));
  }

  void remember(BinarySourceSpec source, Optional<JsonLocation> json) {
    if (original.isEmpty()) {
      original = Optional.of(source);
      location = json;
    }
  }

  void reserve(long bytes) {
    if (reserved < 0) {
      budget.reserveOccurrence(bytes);
      reserved = bytes;
    } else if (reserved != bytes)
      throw new UnexpectedNativeContentException(
          new IllegalStateException("Media materialization length changed"));
  }

  BinarySourceSpec materialize(PictureDataSpec image, ExecutionInputBindings bindings)
      throws IOException {
    try {
      if (image.source() instanceof BinarySourceSpec.InlineBase64 inline) {
        reserve(BinarySourceSpec.decodedSize(inline));
        return inline;
      }
      InputSnapshotLimit limit = readLimit();
      byte[] bytes =
          image.source() instanceof BinarySourceSpec.File file
              ? file(image, file, bindings, limit)
              : standardInput(image, bindings, limit);
      return new BinarySourceSpec.InlineBase64(Base64.getEncoder().encodeToString(bytes));
    } catch (InvalidRequestException failure) {
      if (failure.requestProblem() instanceof FieldValidationProblem problem
          && problem.rule() == FieldValidationMediaRule.RESOURCE_LIMIT)
        budget.poison(problem.operands().getFirst());
      throw failure;
    }
  }

  private byte[] file(
      PictureDataSpec image,
      BinarySourceSpec.File file,
      ExecutionInputBindings bindings,
      InputSnapshotLimit limit)
      throws IOException {
    Path snapshot;
    try {
      snapshot =
          bindings
              .requestPathAccess()
              .materializeReadWithin(
                  file.path(), "picture payload", "gridgrind-media-", ".bin", limit);
    } catch (java.nio.file.InvalidPathException failure) {
      throw new InputSourceReadException(
          "Invalid picture payload path: " + file.path(), "picture payload", file.path(), failure);
    } catch (IOException failure) {
      throw SourceBackedPlanResolver.inputFileFailure(file.path(), "picture payload", failure);
    }
    try {
      long size = Files.size(snapshot);
      limit.require(size);
      reserve(size);
      requireContent(image, size);
      try (var input = Files.newInputStream(snapshot)) {
        byte[] bytes = input.readNBytes(Math.toIntExact(size));
        if (bytes.length != size || input.read() != -1)
          throw new UnexpectedNativeContentException(
              new IllegalStateException("Immutable media snapshot changed"));
        return bytes;
      }
    } catch (IOException failure) {
      throw new InputSourceReadException(
          "Failed to read picture payload file: " + file.path(),
          "picture payload",
          file.path(),
          failure);
    }
  }

  private byte[] standardInput(
      PictureDataSpec image, ExecutionInputBindings bindings, InputSnapshotLimit limit)
      throws IOException {
    var view =
        bindings
            .standardInputSnapshotWithin(limit)
            .orElseThrow(() ->
                new InputSourceUnavailableException(
                    "Standard-input picture payload is unavailable", "picture payload"));
    reserve(view.remaining());
    requireContent(image, view.remaining());
    byte[] bytes = new byte[view.remaining()];
    view.get(bytes);
    return bytes;
  }

  private static void requireContent(PictureDataSpec image, long size) {
    if (size == 0) throw MediaIssues.structure(image.format(), MediaStructureReason.TRUNCATED, 0);
  }

  void reserveBound(
      WorkbookStep authored, WorkbookStep bound, ExecutionInputBindings bindings, int index) {
    Optional<PictureDataSpec> image = image(bound);
    if (image.isEmpty()) return;
    if (original.isEmpty())
      image(authored)
          .ifPresent(value -> {
            String path = "steps[" + index + "].action." + field(authored) + ".source";
            remember(
                value.source(),
                bindings.withInputResolutionScope(path).inputOrigin(value.source()));
          });
    if (!(image.orElseThrow().source() instanceof BinarySourceSpec.InlineBase64 inline))
      throw new UnexpectedNativeContentException(
          new IllegalStateException("Media source was not materialized"));
    reserve(BinarySourceSpec.decodedSize(inline));
  }

  void inspect(PreparedStep prepared) {
    if (!(prepared instanceof PreparedMutationStep mutation)) return;
    switch (mutation.action()) {
      case ResolvedMutationAction.SetPicture action ->
          inspectContent(action.picture().format(), action.picture().imageData());
      case ResolvedMutationAction.SetEmbeddedObject action ->
          inspectContent(
              action.embeddedObject().previewFormat(), action.embeddedObject().previewImage());
      case ResolvedMutationAction.SetSignatureLine action -> {
        var value = action.signatureLine();
        if (value.plainSignature().isPresent())
          inspectContent(
              value.plainSignatureFormat().orElseThrow(), value.plainSignature().orElseThrow());
      }
      default -> {}
    }
  }

  private void inspectContent(
      dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat format,
      dev.erst.gridgrind.core.source.BinaryContent content) {
    if (reserved != content.size())
      throw new UnexpectedNativeContentException(
          new IllegalStateException("Prepared media lacks its exact preallocation reservation"));
    AuthoredImageAdmission.inspect(format, content, budget);
  }

  InputResolutionFailure failure(Exception failure) {
    return original
        .map(source -> InputResolutionFailure.forSource(failure, location, source))
        .orElseGet(() -> InputResolutionFailure.unlocated(failure));
  }

  private static Optional<PictureDataSpec> image(WorkbookStep step) {
    if (!(step instanceof MutationStep mutation)) return Optional.empty();
    return switch (mutation.action()) {
      case DrawingMutationAction.SetPicture action -> Optional.of(action.picture().image());
      case DrawingMutationAction.SetEmbeddedObject action ->
          Optional.of(action.embeddedObject().previewImage());
      case DrawingMutationAction.SetSignatureLine action -> action.signatureLine().plainSignature();
      default -> Optional.empty();
    };
  }

  private static String field(WorkbookStep step) {
    return switch (((MutationStep) step).action()) {
      case DrawingMutationAction.SetPicture _ -> "picture.image";
      case DrawingMutationAction.SetEmbeddedObject _ -> "embeddedObject.previewImage";
      case DrawingMutationAction.SetSignatureLine _ -> "signatureLine.plainSignature";
      default -> throw new IllegalStateException();
    };
  }
}
