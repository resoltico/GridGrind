package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.problem.FieldValidationAddressRule;
import dev.erst.gridgrind.core.problem.FieldValidationBasicRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** One ordered concrete rectangle, with zero-based geometry and canonical A1 presentation. */
public record RangeRef(int firstRow, int lastRow, int firstColumn, int lastColumn)
    implements WireString {
  public RangeRef {
    new CellAddress(firstRow, firstColumn);
    new CellAddress(lastRow, lastColumn);
    if (lastRow < firstRow || lastColumn < firstColumn) {
      throw invalid(
          FieldValidationProblem.detached("range", FieldValidationAddressRule.RANGE_ORDER), null);
    }
  }

  /** Parses one cell or ordered rectangle; single-cell and degenerate rectangles share identity. */
  public RangeRef(String value) {
    this(parse(value));
  }

  private RangeRef(RangeRef range) {
    this(range.firstRow(), range.lastRow(), range.firstColumn(), range.lastColumn());
  }

  /** Parses the concrete geometry of one local A1 rectangle. */
  public static RangeRef parse(String value) {
    Objects.requireNonNull(value, "range must not be null");
    if (value.isBlank()) {
      throw invalid(
          FieldValidationProblem.detached("range", FieldValidationBasicRule.NON_BLANK), null);
    }
    String[] endpoints = value.split(":", -1);
    if (endpoints.length > 2) {
      throw invalid(
          FieldValidationProblem.detached(
              "range", FieldValidationAddressRule.RANGE_RECTANGULAR_SYNTAX),
          null);
    }
    CellAddress first = endpoint(endpoints[0]);
    CellAddress last = endpoints.length == 1 ? first : endpoint(endpoints[1]);
    return new RangeRef(first.rowIndex(), last.rowIndex(), first.columnIndex(), last.columnIndex());
  }

  private static CellAddress endpoint(String value) {
    try {
      return new CellAddress(value);
    } catch (InvalidRequestException failure) {
      var problem = (FieldValidationProblem) failure.requestProblem();
      throw invalid(
          new FieldValidationProblem("range", Optional.empty(), problem.rule(), problem.operands()),
          failure);
    }
  }

  /** Returns the number of rows in this rectangle. */
  public int rowCount() {
    return lastRow - firstRow + 1;
  }

  /** Returns the number of columns in this rectangle. */
  public int columnCount() {
    return lastColumn - firstColumn + 1;
  }

  /** Returns the exact cell count without integer overflow. */
  public long cellCount() {
    return (long) rowCount() * columnCount();
  }

  /** Returns whether this rectangle contains one concrete cell. */
  public boolean contains(CellAddress address) {
    return address.rowIndex() >= firstRow
        && address.rowIndex() <= lastRow
        && address.columnIndex() >= firstColumn
        && address.columnIndex() <= lastColumn;
  }

  /** Returns whether this rectangle and another share at least one cell. */
  public boolean intersects(RangeRef other) {
    Objects.requireNonNull(other, "other must not be null");
    return firstRow <= other.lastRow
        && lastRow >= other.firstRow
        && firstColumn <= other.lastColumn
        && lastColumn >= other.firstColumn;
  }

  @Override
  public String value() {
    String first = new CellAddress(firstRow, firstColumn).value();
    return firstRow == lastRow && firstColumn == lastColumn
        ? first
        : first + ":" + new CellAddress(lastRow, lastColumn).value();
  }

  private static InvalidRequestException invalid(
      FieldValidationProblem problem, @Nullable Throwable cause) {
    return new InvalidRequestException(
        new FieldValidationProblem(
            problem.fieldName(), Optional.empty(), problem.rule(), problem.operands()),
        Optional.empty(),
        Optional.empty(),
        Optional.empty(),
        cause);
  }

  @Override
  public String toString() {
    return value();
  }
}
