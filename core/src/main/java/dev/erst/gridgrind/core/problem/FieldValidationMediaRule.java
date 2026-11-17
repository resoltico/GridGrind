package dev.erst.gridgrind.core.problem;

/** Known authored-image structural, resource and profile refusals. */
public enum FieldValidationMediaRule implements FieldValidationRule {
  FORMAT_STRUCTURE,
  RESOURCE_LIMIT,
  PROFILE_UNVERIFIABLE;

  @Override
  public String message(FieldValidationProblem problem) {
    return switch (this) {
      case FORMAT_STRUCTURE ->
          problem.fieldName()
              + " has invalid "
              + problem.operand(0)
              + " image structure: "
              + MediaStructureReason.valueOf(problem.operand(1)).description()
              + " at byte "
              + problem.operand(2)
              + ".";
      case RESOURCE_LIMIT ->
          problem.fieldName() + " exceeds media limit " + problem.operand(0) + ".";
      case PROFILE_UNVERIFIABLE ->
          problem.fieldName()
              + " declares an unqualified "
              + problem.operand(0)
              + " image profile: "
              + problem.operand(1)
              + ".";
    };
  }

  @Override
  public String resolution(FieldValidationProblem problem) {
    return switch (this) {
      case FORMAT_STRUCTURE ->
          "Provide a structurally complete image matching the declared format.";
      case RESOURCE_LIMIT -> "Reduce the authored media to fit the declared core limit.";
      case PROFILE_UNVERIFIABLE ->
          "Provide an image in an admitted format profile; existing opaque media can still be read"
              + " or copied.";
    };
  }
}
