package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;
import java.util.Objects;

/** Rectangular cell window returned in sparse or dense form. */
@JsonTypeInfo(
    use = JsonTypeInfo.Id.NAME,
    include = JsonTypeInfo.As.EXISTING_PROPERTY,
    property = "shape")
@JsonSubTypes({
  @JsonSubTypes.Type(value = WindowReport.Sparse.class, name = "SPARSE"),
  @JsonSubTypes.Type(value = WindowReport.Dense.class, name = "DENSE")
})
public sealed interface WindowReport permits WindowReport.Sparse, WindowReport.Dense {
  /** Returns the published window shape discriminator. */
  String shape();

  /** Returns the sheet owning the returned window. */
  String sheetName();

  /** Returns the top-left A1 address anchoring this window. */
  String topLeftAddress();

  /** Returns the requested rectangular dimensions of this window. */
  WindowDimensionsReport dimensions();

  /** Sparse window shape that omits blank cells entirely. */
  record Sparse(
      String sheetName,
      String topLeftAddress,
      WindowDimensionsReport dimensions,
      List<CellReport> populatedCells)
      implements WindowReport {
    public Sparse {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      Objects.requireNonNull(dimensions, "dimensions must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (topLeftAddress.isBlank()) {
        throw new IllegalArgumentException("topLeftAddress must not be blank");
      }
      populatedCells = GridGrindResponseSupport.copyValues(populatedCells, "populatedCells");
      for (CellReport cell : populatedCells) {
        if (cell instanceof CellReport.BlankReport) {
          throw new IllegalArgumentException("populatedCells must not contain blank cells");
        }
      }
    }

    @Override
    @JsonProperty
    public String shape() {
      return "SPARSE";
    }
  }

  /** Dense window shape that retains the explicit row grid, including blanks. */
  record Dense(
      String sheetName,
      String topLeftAddress,
      WindowDimensionsReport dimensions,
      List<WindowRowReport> rows)
      implements WindowReport {
    public Dense {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      Objects.requireNonNull(dimensions, "dimensions must not be null");
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetName must not be blank");
      }
      if (topLeftAddress.isBlank()) {
        throw new IllegalArgumentException("topLeftAddress must not be blank");
      }
      rows = GridGrindResponseSupport.copyValues(rows, "rows");
    }

    @Override
    @JsonProperty
    public String shape() {
      return "DENSE";
    }
  }
}
