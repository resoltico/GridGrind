package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Resolves the strongest retained path for one independently bound request fragment. */
final class RequestBindingPathSupport {
  private RequestBindingPathSupport() {}

  static Optional<String> preciseInnerPath(
      RequestJsonNode fragment, Optional<String> jacksonPath, Optional<String> requestProblemPath) {
    if (jacksonPath
        .filter(
            path ->
                GridGrindJsonPathSupport.pathExists(
                    RequestFragmentBinder.toJsonNode(fragment), path))
        .isPresent()) {
      return GridGrindJsonPathSupport.qualifyPath(jacksonPath, requestProblemPath);
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

  private static String formulaSourcePath(RequestJsonNode fragment, String jacksonPath) {
    if ("value".equals(jacksonPath)
        && GridGrindJsonPathSupport.pathExists(
            RequestFragmentBinder.toJsonNode(fragment), "action.value.source.text")) {
      return "action.value.source.text";
    }
    return jacksonPath;
  }
}
