package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.MediaStructureReason;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import java.util.Optional;

/** Constructs only established authored-media refusals with core-owned wording. */
public final class MediaIssues {
  private MediaIssues() {}

  /** Refuses a concrete malformed or incomplete declared image structure. */
  public static InvalidRequestException structure(
      ExcelPictureFormat format, MediaStructureReason reason, long offset) {
    return problem(
        FieldValidationMediaRule.FORMAT_STRUCTURE,
        format.name(),
        reason.name(),
        Long.toString(offset));
  }

  /** Refuses work exceeding a visible core-owned media limit. */
  public static InvalidRequestException limit(String id) {
    return problem(FieldValidationMediaRule.RESOURCE_LIMIT, id);
  }

  /** Refuses a framed profile whose coding grammar is not qualified. */
  public static InvalidRequestException profile(ExcelPictureFormat format, String profile) {
    return problem(FieldValidationMediaRule.PROFILE_UNVERIFIABLE, format.name(), profile);
  }

  static InvalidRequestException coding(
      ExcelPictureFormat format, long offset, java.util.zip.DataFormatException cause) {
    return new InvalidRequestException(
        FieldValidationProblem.detached(
            "image",
            FieldValidationMediaRule.FORMAT_STRUCTURE,
            format.name(),
            MediaStructureReason.CODE.name(),
            Long.toString(offset)),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        cause);
  }

  private static InvalidRequestException problem(
      FieldValidationMediaRule rule, String... operands) {
    return new InvalidRequestException(
        FieldValidationProblem.detached("image", rule, operands),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        null);
  }
}
