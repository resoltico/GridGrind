package dev.erst.gridgrind.core.vocabulary;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

/** Preserves exact text and field-specific failures at the common primitive boundary. */
class TextValidationTest {
  @ParameterizedTest
  @ValueSource(strings = {"", " ", "\t\n", "\u2003"})
  void rejectsUnicodeBlankText(String value) {
    assertEquals(
        "label must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> TextValidation.requireNonBlank(value, "label"))
            .getMessage());
  }

  @Test
  void retainsExactNonblankTextAndNullFailure() {
    assertEquals("  Label  ", TextValidation.requireNonBlank("  Label  ", "label"));
    assertEquals("\u200B", TextValidation.requireNonBlank("\u200B", "label"));
    assertEquals(
        "label must not be null",
        assertThrows(
                NullPointerException.class, () -> TextValidation.requireNonBlank(null, "label"))
            .getMessage());
  }
}
