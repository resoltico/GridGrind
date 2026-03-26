package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import java.math.BigDecimal;

/** JSON-friendly typed font height input used by style patches. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = FontHeightInput.Points.class, name = "POINTS"),
  @JsonSubTypes.Type(value = FontHeightInput.Twips.class, name = "TWIPS")
})
public sealed interface FontHeightInput {

  /** Converts this protocol font-height value into the workbook-core representation. */
  ExcelFontHeight toExcelFontHeight();

  /** Font height expressed in point units, such as {@code 11} or {@code 11.5}. */
  record Points(BigDecimal points) implements FontHeightInput {
    public Points {
      ExcelFontHeight.fromPoints(points);
    }

    @Override
    public ExcelFontHeight toExcelFontHeight() {
      return ExcelFontHeight.fromPoints(points);
    }
  }

  /** Font height expressed in exact twips, where one point equals twenty twips. */
  record Twips(int twips) implements FontHeightInput {
    public Twips {
      new ExcelFontHeight(twips);
    }

    @Override
    public ExcelFontHeight toExcelFontHeight() {
      return new ExcelFontHeight(twips);
    }
  }
}
