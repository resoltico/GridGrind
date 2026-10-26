package dev.erst.gridgrind.core.model.layout;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.result.WorkbookResultSupport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.List;
import java.util.Objects;

/** Rectangular cell window returned in sparse or dense form. */
@Union(discriminator = "shape", catalog = true, displayOrder = 3)
public sealed interface WindowReport
    permits dev.erst.gridgrind.core.model.layout.WindowReport.Sparse,
        dev.erst.gridgrind.core.model.layout.WindowReport.Dense {
  /** Returns the published window shape discriminator. */
  String shape();

  /** Returns the sheet owning the returned window. */
  SheetLabel sheetName();

  /** Returns the top-left A1 address anchoring this window. */
  CellAddress topLeftAddress();

  /** Returns the requested rectangular dimensions of this window. */
  WindowDimensionsReport dimensions();

  /** Sparse window shape that omits blank cells entirely. */
  @Variant(
      id = "SPARSE",
      summary = "Sparse window omitting blank cells and publishing populatedCells only.")
  record Sparse(
      SheetLabel sheetName,
      CellAddress topLeftAddress,
      WindowDimensionsReport dimensions,
      List<CellReport> populatedCells)
      implements WindowReport {
    public Sparse {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      Objects.requireNonNull(dimensions, "dimensions must not be null");
      populatedCells = WorkbookResultSupport.copyValues(populatedCells, "populatedCells");
      for (CellReport cell : populatedCells) {
        if (cell instanceof CellReport.BlankReport) {
          throw new IllegalArgumentException("populatedCells must not contain blank cells");
        }
      }
    }

    @Override
    public String shape() {
      return "SPARSE";
    }
  }

  /** Dense window shape that retains the explicit row grid, including blanks. */
  @Variant(
      id = "DENSE",
      summary = "Dense window preserving explicit row structure including blank cells.")
  record Dense(
      SheetLabel sheetName,
      CellAddress topLeftAddress,
      WindowDimensionsReport dimensions,
      List<WindowRowReport> rows)
      implements WindowReport {
    public Dense {
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
      Objects.requireNonNull(dimensions, "dimensions must not be null");
      rows = WorkbookResultSupport.copyValues(rows, "rows");
    }

    @Override
    public String shape() {
      return "DENSE";
    }
  }
}
