package dev.erst.gridgrind.excel;

import java.util.EnumSet;
import java.util.Objects;
import java.util.Set;

/** Internal normalized projection for cell-returning workbook reads. */
public record ExcelCellReadProjection(Set<ExcelCellReadFacet> facets) {
  private static final Set<ExcelCellReadFacet> DEFAULT_FACETS = Set.of(ExcelCellReadFacet.VALUE);

  public ExcelCellReadProjection {
    Objects.requireNonNull(facets, "facets must not be null");
    if (facets.isEmpty()) {
      throw new IllegalArgumentException("facets must not be empty");
    }
    Set<ExcelCellReadFacet> copy = EnumSet.noneOf(ExcelCellReadFacet.class);
    for (ExcelCellReadFacet facet : facets) {
      copy.add(Objects.requireNonNull(facet, "facets must not contain nulls"));
    }
    facets = Set.copyOf(copy);
  }

  /** Returns the default internal projection used when a workbook read omits facet selection. */
  public static ExcelCellReadProjection defaults() {
    return new ExcelCellReadProjection(DEFAULT_FACETS);
  }

  /** Returns whether the internal projection requests the provided read facet. */
  public boolean includes(ExcelCellReadFacet facet) {
    return facets.contains(Objects.requireNonNull(facet, "facet must not be null"));
  }
}
