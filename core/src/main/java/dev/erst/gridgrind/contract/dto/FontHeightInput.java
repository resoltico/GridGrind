package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.math.BigDecimal;
import java.util.Objects;

/** JSON-friendly typed font height input used by style patches. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = FontHeightInput.Points.class, name = "POINTS"),
  @JsonSubTypes.Type(value = FontHeightInput.Twips.class, name = "TWIPS")
})
public sealed interface FontHeightInput {

  /** Font height expressed in point units, such as {@code 11} or {@code 11.5}. */
  record Points(BigDecimal points) implements FontHeightInput {
    public Points {
      validatePoints(points);
    }
  }

  /** Font height expressed in exact twips, where one point equals twenty twips. */
  record Twips(int twips) implements FontHeightInput {
    public Twips {
      validateTwips(twips);
    }
  }

  private static void validatePoints(BigDecimal points) {
    Objects.requireNonNull(points, "points must not be null");
    if (points.signum() <= 0) {
      throw new IllegalArgumentException("points must be greater than 0");
    }
    try {
      validateTwips(points.multiply(BigDecimal.valueOf(20)).intValueExact());
    } catch (ArithmeticException exception) {
      throw new IllegalArgumentException(
          "points must resolve exactly to whole twips (1/20 points)", exception);
    }
  }

  private static void validateTwips(int twips) {
    if (twips <= 0 || twips > Short.MAX_VALUE) {
      throw new IllegalArgumentException(
          "twips must be between 1 and " + Short.MAX_VALUE + " (inclusive)");
    }
  }
}
