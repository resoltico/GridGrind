package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.core.problem.RequestMalformedScalar;
import dev.erst.gridgrind.core.problem.RequestNumberNotRepresentable;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

/** Checks Java numeric-token fidelity and temporal syntax without a trial Jackson binding. */
final class RequestScalarTokens {
  private RequestScalarTokens() {}

  static void validate(
      RequestJsonNode node,
      Class<?> type,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (RequestJavaTypes.isTemporalType(type)) {
      temporal(node, type, path, offset, problems);
      return;
    }
    if (!(node instanceof RequestJsonNumber number)) {
      problems.add(new RequestMalformedScalar(path, "a JSON number", offset));
      return;
    }
    if ((type == double.class || type == Double.class || type == float.class || type == Float.class)
        && !exact(number.value(), type)) {
      problems.add(new RequestNumberNotRepresentable(path, number.value(), offset));
      return;
    }
    try {
      numeric(number.value(), type);
    } catch (NumberFormatException | ArithmeticException invalid) {
      problems.add(
          new RequestMalformedScalar(path, RequestJavaTypes.numericExpectation(type), offset));
    }
  }

  private static void numeric(String token, Class<?> type) {
    if (type == byte.class || type == Byte.class) {
      Byte.parseByte(token);
    } else if (type == short.class || type == Short.class) {
      Short.parseShort(token);
    } else if (type == int.class || type == Integer.class) {
      Integer.parseInt(token);
    } else if (type == long.class || type == Long.class) {
      Long.parseLong(token);
    } else if (type == BigInteger.class) {
      new BigInteger(token);
    } else if (type == BigDecimal.class) {
      new BigDecimal(token);
    } else if (!(type == double.class
        || type == Double.class
        || type == float.class
        || type == Float.class)) {
      throw new IllegalStateException("Unsupported numeric token type: " + type.getName());
    }
  }

  private static boolean exact(String token, Class<?> type) {
    try {
      BigDecimal authored = new BigDecimal(token);
      if (type == float.class || type == Float.class) {
        float value = Float.parseFloat(token);
        return Float.isFinite(value) && authored.compareTo(BigDecimal.valueOf((double) value)) == 0;
      }
      double value = Double.parseDouble(token);
      return Double.isFinite(value) && authored.compareTo(BigDecimal.valueOf(value)) == 0;
    } catch (NumberFormatException invalid) {
      return false;
    }
  }

  private static void temporal(
      RequestJsonNode node,
      Class<?> type,
      String path,
      long offset,
      List<RequestStructuralProblem> problems) {
    if (!(node instanceof RequestJsonString string)) {
      problems.add(new RequestMalformedScalar(path, "a JSON string", offset));
      return;
    }
    try {
      if (type == LocalDate.class) {
        LocalDate.parse(string.value());
      } else if (type == LocalDateTime.class) {
        LocalDateTime.parse(string.value());
      } else {
        throw new IllegalStateException("Unsupported temporal token type: " + type.getName());
      }
    } catch (java.time.format.DateTimeParseException invalid) {
      problems.add(
          new RequestMalformedScalar(path, RequestJavaTypes.temporalExpectation(type), offset));
    }
  }
}
