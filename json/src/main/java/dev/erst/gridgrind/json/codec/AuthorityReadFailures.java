package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.problem.AuthorityDocumentIssue;
import dev.erst.gridgrind.core.problem.InvalidAuthorityException;
import dev.erst.gridgrind.core.problem.NumberNotRepresentableException;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.core.problem.RequestProblemDescriptor;
import dev.erst.gridgrind.core.problem.RequestProblemSource;
import java.lang.reflect.InvocationTargetException;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Pattern;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JsonParser;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.core.exc.StreamReadException;
import tools.jackson.databind.exc.InvalidDefinitionException;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.exc.UnrecognizedPropertyException;
import tools.jackson.databind.exc.ValueInstantiationException;

/** Expected-first classification for the strict authority document reader alone. */
final class AuthorityReadFailures {
  private static final Pattern FIELD = Pattern.compile("[A-Za-z_][A-Za-z0-9_]{0,127}");

  private AuthorityReadFailures() {}

  static RuntimeException translate(RuntimeException failure, JsonParser parser) {
    if (failure instanceof InvalidAuthorityException known) return known;
    if (failure instanceof InvalidDefinitionException)
      return new UnexpectedAuthorityReadException(failure);
    if (failure instanceof tools.jackson.core.exc.InputCoercionException value)
      return refusal(AuthorityDocumentIssue.VALUE, path(value), cursor(value, parser), failure);
    if (failure instanceof StreamReadException syntax)
      return refusal(AuthorityDocumentIssue.SYNTAX, path(syntax), cursor(syntax, parser), failure);
    if (failure instanceof MismatchedInputException shape)
      return refusal(AuthorityDocumentIssue.SHAPE, path(shape), cursor(shape, parser), failure);
    if (failure instanceof ValueInstantiationException construction) {
      Throwable cause = construction.getCause();
      if (cause instanceof InvocationTargetException invocation)
        cause = invocation.getTargetException();
      if (cause instanceof Error fatal) throw fatal;
      if (cause instanceof RequestProblemSource declared) {
        boolean scalar =
            construction.getType() != null
                && dev.erst.gridgrind.core.vocabulary.WireScalar.isValueType(
                    construction.getType().getRawClass());
        return declaredFailure(declared, path(construction), scalar, failure);
      }
      if (cause instanceof NumberNotRepresentableException)
        return refusal(
            AuthorityDocumentIssue.VALUE,
            path(construction),
            cursor(construction, parser),
            failure);
      return new UnexpectedAuthorityReadException(failure);
    }
    if (failure.getClass() == tools.jackson.databind.DatabindException.class
        && failure.getCause() instanceof RequestProblemSource declared) {
      return declaredFailure(declared, path((JacksonException) failure), false, failure);
    }
    if (failure instanceof RequestProblemSource declared) {
      return declaredFailure(declared, Optional.empty(), false, failure);
    }
    if (failure instanceof NumberNotRepresentableException)
      return refusal(
          AuthorityDocumentIssue.VALUE, Optional.empty(), parser.currentTokenLocation(), failure);
    return new UnexpectedAuthorityReadException(failure);
  }

  private static InvalidAuthorityException declaredFailure(
      RequestProblemSource source, Optional<String> parent, boolean scalar, Throwable failure) {
    Optional<String> local =
        source instanceof dev.erst.gridgrind.core.problem.PayloadException payload
            ? payload.jsonPath().or(source.requestProblem()::jsonPath)
            : source.requestProblem().jsonPath();
    Optional<String> field = scalar ? parent : combine(parent, local);
    return new InvalidAuthorityException(
        source.requestProblem() instanceof RequestProblemDescriptor.Shape
            ? AuthorityDocumentIssue.SHAPE
            : AuthorityDocumentIssue.VALUE,
        field.map(JsonLocation::pathOnly).orElseGet(JsonLocation::unavailable),
        failure);
  }

  static InvalidAuthorityException rootRefusal(JsonParser parser) {
    return refusal(
        AuthorityDocumentIssue.SHAPE, Optional.empty(), parser.currentTokenLocation(), null);
  }

  private static InvalidAuthorityException refusal(
      AuthorityDocumentIssue issue,
      Optional<String> path,
      @Nullable TokenStreamLocation cursor,
      @Nullable Throwable failure) {
    long offset = cursor == null ? -1 : cursor.getByteOffset();
    int line = cursor == null ? -1 : cursor.getLineNr();
    int column = cursor == null ? -1 : cursor.getColumnNr();
    JsonLocation location;
    if (path.isPresent()) {
      location =
          offset >= 0
              ? JsonLocation.pathAtByteOffset(path.orElseThrow(), offset)
              : line > 0 && column > 0
                  ? JsonLocation.located(path.orElseThrow(), line, column)
                  : JsonLocation.pathOnly(path.orElseThrow());
    } else {
      location =
          offset >= 0
              ? line > 0 && column > 0
                  ? JsonLocation.byteOffsetLineColumn(offset, line, column)
                  : JsonLocation.byteOffset(offset)
              : line > 0 && column > 0
                  ? JsonLocation.lineColumn(line, column)
                  : JsonLocation.unavailable();
    }
    return new InvalidAuthorityException(issue, location, failure);
  }

  private static @Nullable TokenStreamLocation cursor(JacksonException failure, JsonParser parser) {
    TokenStreamLocation own = failure.getLocation();
    return own != null && (own.getByteOffset() >= 0 || own.getLineNr() > 0)
        ? own
        : parser.currentTokenLocation();
  }

  private static Optional<String> path(JacksonException failure) {
    StringBuilder value = new StringBuilder();
    var references = failure.getPath();
    boolean unknownProperty =
        failure instanceof UnrecognizedPropertyException
            || (failure.getCause()
                    instanceof dev.erst.gridgrind.core.problem.InvalidRequestShapeException shape
                && shape.requestProblem() instanceof dev.erst.gridgrind.core.problem.UnknownField);
    int size = references.size() - (unknownProperty ? 1 : 0);
    for (int index = 0; index < size; index++) {
      var reference = references.get(index);
      Object from = reference.from();
      if (from instanceof Map<?, ?>
          || (from instanceof Class<?> type && Map.class.isAssignableFrom(type))) break;
      String name = reference.getPropertyName();
      if (name != null) {
        if (!FIELD.matcher(name).matches()) break;
        if (!value.isEmpty()) value.append('.');
        value.append(name);
      } else if (reference.getIndex() >= 0)
        value.append('[').append(reference.getIndex()).append(']');
    }
    return value.isEmpty() ? Optional.empty() : Optional.of(value.toString());
  }

  private static Optional<String> combine(Optional<String> parent, Optional<String> child) {
    if (child.isEmpty()) return parent;
    String local = child.orElseThrow();
    // Logical model paths come from declared rules, never framework exception messages.
    if (parent.isEmpty()
        || parent.orElseThrow().equals(local)
        || parent.orElseThrow().endsWith('.' + local)) return parent.or(() -> child);
    return Optional.of(parent.orElseThrow() + '.' + local);
  }
}
