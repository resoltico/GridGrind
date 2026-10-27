package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Shared read projection used by cell-returning inspection queries. */
@CatalogType(
    id = "CellReadProjection",
    summary =
        "Shared readback projection that selects which factual cell facets are included on cell-returning inspection surfaces. Facets default to VALUE when omitted by the enclosing query. The facets field publishes enumValueDocs so each facet token explains the response fields it unlocks. Catalog field descriptors publish facet-gated response fields through projectedByFacets so agents can derive the required facet directly from the machine contract. Date-like numeric cells still read back as type=NUMBER until TEMPORAL is requested.",
    displayOrder = 47)
public record CellReadProjection(List<CellReadFacet> facets) {
  private static final List<CellReadFacet> DEFAULT_FACETS = List.of(CellReadFacet.VALUE);

  public CellReadProjection {
    Objects.requireNonNull(facets, "facets must not be null");
    if (facets.isEmpty()) {
      throw new IllegalArgumentException("facets must not be empty");
    }
    Set<CellReadFacet> copy = new LinkedHashSet<>(facets.size());
    for (CellReadFacet facet : facets) {
      copy.add(Objects.requireNonNull(facet, "facets must not contain nulls"));
    }
    facets = List.copyOf(new ArrayList<>(copy));
  }

  /** Returns the default projection used when a caller omits facet selection. */
  public static CellReadProjection defaults() {
    return new CellReadProjection(DEFAULT_FACETS);
  }

  /** Returns a projection containing the provided facets in encounter order without duplicates. */
  public static CellReadProjection of(CellReadFacet first, CellReadFacet... remaining) {
    Objects.requireNonNull(first, "first must not be null");
    Objects.requireNonNull(remaining, "remaining must not be null");
    List<CellReadFacet> values = new ArrayList<>(remaining.length + 1);
    values.add(first);
    for (CellReadFacet facet : remaining) {
      values.add(Objects.requireNonNull(facet, "remaining must not contain nulls"));
    }
    return new CellReadProjection(values);
  }

  /** Returns whether the projection requests the provided facet. */
  public boolean includes(CellReadFacet facet) {
    return facets.contains(Objects.requireNonNull(facet, "facet must not be null"));
  }
}
