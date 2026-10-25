package dev.erst.gridgrind.contract.json;

import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Resolves the strongest retained path for one independently bound request fragment. */
final class RequestBindingPathSupport {
  private static final Pattern LEADING_FIELD_NAME =
      Pattern.compile("^([A-Za-z][A-Za-z0-9]*) must ");

  private RequestBindingPathSupport() {}

  static Optional<String> preciseInnerPath(
      RequestJsonNode fragment, Optional<String> jacksonPath, Optional<String> requestProblemPath) {
    if (jacksonPath
        .filter(
            path ->
                GridGrindJsonPathSupport.pathExists(
                    RequestFragmentBinder.toJsonNode(fragment), path))
        .isPresent()) {
      return jacksonPath;
    }
    return requestProblemPath.or(() -> jacksonPath);
  }

  static Optional<tools.jackson.core.JacksonException> jacksonFailure(Throwable failure) {
    Throwable current = failure;
    while (current != null) {
      if (current instanceof tools.jackson.core.JacksonException jacksonException) {
        return Optional.of(jacksonException);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  static Optional<String> qualifiedFormulaFailurePath(
      RequestJsonNode fragment, String fragmentPath, Throwable failure) {
    Optional<String> nestedPath =
        jacksonFailure(failure)
            .flatMap(
                exception ->
                    GridGrindJsonPayloadMetadataSupport.payloadMetadata(exception).jsonPath());
    return nestedPath
        .map(path -> formulaSourcePath(fragment, path))
        .map(
            path ->
                GridGrindJsonPathSupport.qualifyPath(Optional.of(fragmentPath), Optional.of(path))
                    .orElse(fragmentPath))
        .or(() -> Optional.of(fragmentPath));
  }

  static Optional<String> directChildPathForInvariant(RequestJsonNode fragment, Throwable failure) {
    Matcher matcher = LEADING_FIELD_NAME.matcher(String.valueOf(failure.getMessage()));
    if (!matcher.find()) {
      return Optional.empty();
    }
    String fieldName = matcher.group(1);
    java.util.List<String> paths = new java.util.ArrayList<>();
    collectFieldPaths(fragment, fieldName, "", paths);
    return paths.size() == 1 ? Optional.of(paths.getFirst()) : Optional.empty();
  }

  static Optional<String> inferQualifiedInvariantFieldPath(
      RequestJsonNode root, RequestBindingFailure failure) {
    Matcher matcher = LEADING_FIELD_NAME.matcher(String.valueOf(failure.exception().getMessage()));
    if (!matcher.find()) {
      return Optional.empty();
    }
    String candidate = failure.jsonPath() + "." + matcher.group(1);
    return RequestJsonTokenLocationSupport.byteOffsetAt(root, candidate).isPresent()
        ? Optional.of(candidate)
        : Optional.empty();
  }

  private static void collectFieldPaths(
      RequestJsonNode node, String fieldName, String prefix, java.util.List<String> paths) {
    switch (node) {
      case RequestJsonObject object -> {
        for (RequestJsonMember member : object.members()) {
          String memberPath = prefix.isEmpty() ? member.name() : prefix + "." + member.name();
          if (fieldName.equals(member.name())) {
            paths.add(memberPath);
          }
          collectFieldPaths(member.value(), fieldName, memberPath, paths);
        }
      }
      case RequestJsonArray array -> {
        for (int index = 0; index < array.elements().size(); index++) {
          collectFieldPaths(
              array.elements().get(index), fieldName, prefix + "[" + index + "]", paths);
        }
      }
      default -> {}
    }
  }

  private static String formulaSourcePath(RequestJsonNode fragment, String jacksonPath) {
    if ("value".equals(jacksonPath)
        && GridGrindJsonPathSupport.pathExists(
            RequestFragmentBinder.toJsonNode(fragment), "action.value.source.text")) {
      return "action.value.source.text";
    }
    return jacksonPath;
  }
}
