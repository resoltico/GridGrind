package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Coverage tests for temporal format heuristics and normalization branches. */
class ExcelTemporalFormatSupportCoverageTest {
  @Test
  void observedKindDistinguishesDatesAndDateTimes() {
    assertEquals(Optional.empty(), ExcelTemporalFormatSupport.observedKind("0.00"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.DATE),
        ExcelTemporalFormatSupport.observedKind("m/d/yyyy"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.TIME),
        ExcelTemporalFormatSupport.observedKind("h:mm"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.TIME),
        ExcelTemporalFormatSupport.observedKind("mm:ss"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.DATE_TIME),
        ExcelTemporalFormatSupport.observedKind("m/d/yyyy h:mm"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.TIME),
        ExcelTemporalFormatSupport.observedKind("[h]:mm:ss"));
    assertTrue(
        java.util.EnumSet.allOf(ExcelTemporalFormatSupport.ObservedKind.class)
            .contains(ExcelTemporalFormatSupport.ObservedKind.DATE_TIME));
  }

  @Test
  void normalizeFormatHandlesQuotesEscapesAndBracketSections() {
    assertEquals("m/d/yyyy  h:mm", normalize("m/d/yyyy \"at\" h:mm"));
    assertEquals("mdyyyy", normalize("m\\/d\\/yyyy"));
    assertEquals("h:mm:ss", normalize("[h]:mm:ss"));
    assertEquals("s", normalize("[s]"));
    assertEquals("m/d/yyyy", normalize("[Red]m/d/yyyy"));
    assertEquals("m/d/yyyy [", normalize("m/d/yyyy ["));
    assertEquals("\\", normalize("\\"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.DATE_TIME),
        ExcelTemporalFormatSupport.observedKind("m/d/yyyy AM/PM"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.DATE_TIME),
        ExcelTemporalFormatSupport.observedKind("m/d/yyyy A/P"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.DATE_TIME),
        ExcelTemporalFormatSupport.observedKind("m/d/yyyy h"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.DATE_TIME),
        ExcelTemporalFormatSupport.observedKind("m/d/yyyy s"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.DATE),
        ExcelTemporalFormatSupport.observedKind("m/d/yyyy"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.TIME),
        ExcelTemporalFormatSupport.observedKind("h:mm"));
    assertEquals(
        Optional.of(ExcelTemporalFormatSupport.ObservedKind.TIME),
        ExcelTemporalFormatSupport.observedKind("mm:ss"));
  }

  @Test
  void monthTokenDetectionSeparatesDateAndMinuteContexts() {
    assertTrue(containsDateFields("yyyy"));
    assertTrue(containsDateFields("dd"));
    assertTrue(containsDateFields("mmmm"));
    assertFalse(containsDateFields("h:mm"));
    assertTrue(containsMonthDateToken("mmmm"));
    assertTrue(containsMonthDateToken("m/d"));
    assertFalse(containsMonthDateToken("h:mm"));
    assertFalse(containsMonthDateToken("mm:ss"));
    assertTrue(isMinuteToken("h mm", 2, 4));
    assertTrue(isMinuteToken("mm : ss", 0, 2));
    assertTrue(isMinuteToken("mm ss", 0, 2));
    assertEquals('\0', previousSignificant("mm", -1));
    assertEquals('\0', nextSignificant("mm", 2));
  }

  private static String normalize(String numberFormat) {
    try {
      return (String)
          temporalMethod("normalizeFormat", MethodType.methodType(String.class, String.class))
              .invoke(numberFormat);
    } catch (Throwable throwable) {
      throw new AssertionError(throwable);
    }
  }

  private static boolean containsMonthDateToken(String normalizedFormat) {
    try {
      return (boolean)
          temporalMethod(
                  "containsMonthDateToken", MethodType.methodType(boolean.class, String.class))
              .invoke(normalizedFormat);
    } catch (Throwable throwable) {
      throw new AssertionError(throwable);
    }
  }

  private static boolean containsDateFields(String normalizedFormat) {
    try {
      return (boolean)
          temporalMethod("containsDateFields", MethodType.methodType(boolean.class, String.class))
              .invoke(normalizedFormat);
    } catch (Throwable throwable) {
      throw new AssertionError(throwable);
    }
  }

  private static boolean isMinuteToken(
      String normalizedFormat, int startInclusive, int endExclusive) {
    try {
      return (boolean)
          temporalMethod(
                  "isMinuteToken",
                  MethodType.methodType(boolean.class, String.class, int.class, int.class))
              .invoke(normalizedFormat, startInclusive, endExclusive);
    } catch (Throwable throwable) {
      throw new AssertionError(throwable);
    }
  }

  private static char previousSignificant(String text, int cursor) {
    try {
      return (char)
          temporalMethod(
                  "previousSignificant", MethodType.methodType(char.class, String.class, int.class))
              .invoke(text, cursor);
    } catch (Throwable throwable) {
      throw new AssertionError(throwable);
    }
  }

  private static char nextSignificant(String text, int cursor) {
    try {
      return (char)
          temporalMethod(
                  "nextSignificant", MethodType.methodType(char.class, String.class, int.class))
              .invoke(text, cursor);
    } catch (Throwable throwable) {
      throw new AssertionError(throwable);
    }
  }

  private static MethodHandle temporalMethod(String name, MethodType type) {
    try {
      return MethodHandles.privateLookupIn(ExcelTemporalFormatSupport.class, MethodHandles.lookup())
          .findStatic(ExcelTemporalFormatSupport.class, name, type);
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError(exception.getMessage(), exception);
    }
  }
}
