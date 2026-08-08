package dev.erst.gridgrind.contract.json;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Map;
import java.util.Objects;

/** Classifies creator component types for structural validation. */
final class RequestTypeSupport {
  private static final Map<Class<?>, String> NUMERIC_EXPECTATIONS =
      Map.ofEntries(
          Map.entry(byte.class, "a JSON integer between -128 and 127"),
          Map.entry(Byte.class, "a JSON integer between -128 and 127"),
          Map.entry(short.class, "a JSON integer between -32768 and 32767"),
          Map.entry(Short.class, "a JSON integer between -32768 and 32767"),
          Map.entry(int.class, "a JSON integer between -2147483648 and 2147483647"),
          Map.entry(Integer.class, "a JSON integer between -2147483648 and 2147483647"),
          Map.entry(
              long.class, "a JSON integer between -9223372036854775808 and 9223372036854775807"),
          Map.entry(
              Long.class, "a JSON integer between -9223372036854775808 and 9223372036854775807"),
          Map.entry(float.class, "a finite JSON number representable as a 32-bit float"),
          Map.entry(Float.class, "a finite JSON number representable as a 32-bit float"),
          Map.entry(double.class, "a finite JSON number representable as a 64-bit double"),
          Map.entry(Double.class, "a finite JSON number representable as a 64-bit double"),
          Map.entry(BigInteger.class, "a JSON integer"),
          Map.entry(BigDecimal.class, "a JSON decimal number"));
  private static final Map<Class<?>, String> TEMPORAL_EXPECTATIONS =
      Map.of(
          LocalDate.class, "an ISO-8601 calendar date",
          LocalDateTime.class, "an ISO-8601 local date-time");

  private RequestTypeSupport() {}

  static Class<?> rawType(Type type) {
    return switch (type) {
      case Class<?> rawClass -> rawClass;
      case ParameterizedType parameterized -> rawType(parameterized.getRawType());
      default -> Object.class;
    };
  }

  static Type typeArgument(Type type) {
    if (type instanceof ParameterizedType parameterized
        && parameterized.getActualTypeArguments().length == 1) {
      return parameterized.getActualTypeArguments()[0];
    }
    return Object.class;
  }

  static boolean isNumberType(Class<?> type) {
    return NUMERIC_EXPECTATIONS.containsKey(type);
  }

  static String numericExpectation(Class<?> type) {
    return Objects.requireNonNull(
        NUMERIC_EXPECTATIONS.get(type), "numeric type must be registered before validation");
  }

  static boolean isTemporalType(Class<?> type) {
    return TEMPORAL_EXPECTATIONS.containsKey(type);
  }

  static String temporalExpectation(Class<?> type) {
    return Objects.requireNonNull(
        TEMPORAL_EXPECTATIONS.get(type), "temporal type must be registered before validation");
  }
}
