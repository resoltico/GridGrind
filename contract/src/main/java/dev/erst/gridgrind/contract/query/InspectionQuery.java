package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.WorkbookOperationContracts;
import dev.erst.gridgrind.excel.foundation.ExcelReadLimits;
import java.util.Objects;

/** Ordered post-mutation inspection queries that introspect or analyze workbook state. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
public sealed interface InspectionQuery
    permits InspectionQuery.Introspection, InspectionQuery.Surface, InspectionQuery.Analysis {

  /**
   * Maximum number of factual cells permitted in one cell-returning read surface. Exact-address
   * reads enforce this as an address-count cap, while rectangular reads enforce it on {@code
   * rowCount * columnCount}. Requests exceeding this limit are rejected during plan validation to
   * prevent out-of-memory failures during serialization of large cell grids. See
   * docs/LIMITATIONS.md LIM-001.
   */
  int MAX_READ_CELLS = ExcelReadLimits.MAX_READ_CELLS; // LIM-001

  /** Marker for raw workbook-fact queries with no higher-level interpretation. */
  sealed interface Introspection extends InspectionQuery
      permits WorkbookIntrospectionQuery,
          SheetIntrospectionQuery,
          WorkbookAssetIntrospectionQuery {}

  /** Marker for derived factual surface-summary queries that stop short of health analysis. */
  sealed interface Surface extends InspectionQuery permits InspectionSurfaceQuery {}

  /** Marker for derived workbook analysis queries. */
  sealed interface Analysis extends InspectionQuery permits InspectionAnalysisQuery {}

  /** Returns the stable SCREAMING_SNAKE_CASE discriminator for one inspection query. */
  default String queryType() {
    return GridGrindProtocolTypeNames.inspectionQueryTypeName(
        getClass().asSubclass(InspectionQuery.class));
  }

  /** Returns the selector types accepted by one concrete inspection query instance. */
  static Class<? extends Selector>[] allowedTargetTypes(InspectionQuery query) {
    Objects.requireNonNull(query, "query must not be null");
    return WorkbookOperationContracts.targetSelectorsFor(query);
  }

  /** Returns the selector types accepted by one concrete inspection query type. */
  static Class<? extends Selector>[] allowedTargetTypesForType(
      Class<? extends InspectionQuery> queryType) {
    Objects.requireNonNull(queryType, "queryType must not be null");
    if (!queryType.isRecord()) {
      throw noTargetTypeMapping(queryType);
    }
    try {
      return WorkbookOperationContracts.staticTargetSelectorsFor(queryType);
    } catch (IllegalArgumentException | IllegalStateException exception) {
      throw noTargetTypeMapping(queryType, exception);
    }
  }

  private static IllegalArgumentException noTargetTypeMapping(
      Class<? extends InspectionQuery> queryType) {
    return new IllegalArgumentException(
        "No target-type mapping configured for query class " + queryType.getName());
  }

  private static IllegalArgumentException noTargetTypeMapping(
      Class<? extends InspectionQuery> queryType, RuntimeException cause) {
    return new IllegalArgumentException(
        "No target-type mapping configured for query class " + queryType.getName(), cause);
  }
}
