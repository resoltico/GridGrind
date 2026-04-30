package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;

/** Requested chart title definition. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ChartTitleInput.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = ChartTitleInput.Text.class, name = "TEXT"),
  @JsonSubTypes.Type(value = ChartTitleInput.Formula.class, name = "FORMULA")
})
public sealed interface ChartTitleInput
    permits ChartTitleInput.None, ChartTitleInput.Text, ChartTitleInput.Formula {
  /** Remove any chart or series title. */
  record None() implements ChartTitleInput {}

  /** Use a static text title. */
  record Text(TextSourceInput source) implements ChartTitleInput {
    public Text {
      Objects.requireNonNull(source, "source must not be null");
      if (source instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
        throw new IllegalArgumentException("source must not be blank");
      }
    }
  }

  /** Bind the title to one workbook formula. */
  record Formula(String formula) implements ChartTitleInput {
    public Formula {
      formula = ChartInput.requireNonBlank(formula, "formula");
    }
  }
}
