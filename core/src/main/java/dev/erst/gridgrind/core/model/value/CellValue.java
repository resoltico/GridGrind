package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.Objects;

/** Resolved effective scalar value shared by cell observations and exact-value expectations. */
@Union(discriminator = "type", catalog = true, displayOrder = 0)
public sealed interface CellValue
    permits CellValue.Blank,
        CellValue.Text,
        CellValue.NumberValue,
        CellValue.BooleanValue,
        CellValue.ErrorValue {
  /** Returns this scalar's metadata-owned wire identity. */
  default String type() {
    return dev.erst.gridgrind.core.catalog.ProtocolTypeMetadataSupport.requiredTypeId(getClass());
  }

  /** An effective blank value. */
  @Variant(id = "BLANK", summary = "An effective blank cell value.")
  record Blank() implements CellValue {}

  /** Exact effective text. */
  @Variant(id = "TEXT", summary = "One exact effective text value.")
  record Text(String value) implements CellValue {
    public Text {
      Objects.requireNonNull(value, "value must not be null");
    }
  }

  /** Effective numeric data, including nonfinite values observed in external workbooks. */
  @Variant(
      id = "NUMBER",
      summary =
          "One effective numeric value. Nonfinite observed values retain their JSON string representation; authored writes and numeric expectations must be finite.")
  record NumberValue(double value) implements CellValue {}

  /** Exact effective boolean data. */
  @Variant(id = "BOOLEAN", summary = "One exact effective boolean value.")
  record BooleanValue(boolean value) implements CellValue {}

  /** Reported Excel error data, including evaluation-only states. */
  @Variant(
      id = "ERROR",
      summary =
          "One GridGrind-owned effective error literal, including reported evaluation-only states.")
  record ErrorValue(String value) implements CellValue {
    public ErrorValue {
      value = CellErrorLiteralValidation.requireReportedErrorLiteral(value, "value");
    }
  }
}
