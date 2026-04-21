package dev.erst.gridgrind.excel;

import java.lang.invoke.MethodHandles;
import java.lang.invoke.VarHandle;
import java.util.List;
import java.util.Objects;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;

/**
 * Isolates POI fill-registry access, including the private mutable registry needed for distinct
 * gradient fills.
 *
 * <p>POI's public {@code StylesTable.putFill(...)} path deduplicates through {@link
 * XSSFCellFill#equals(Object)}, which ignores gradient geometry and can alias distinct linear and
 * path gradients. GridGrind therefore appends new gradient fills through the private mutable
 * registry while still reading through POI's public view.
 */
final class StylesTableFillRegistryAccess {
  private static final VarHandle FILLS_FIELD = requireFillsField(MethodHandles.lookup());

  static StylesTableFillRegistryAccess poiApi() {
    return new StylesTableFillRegistryAccess();
  }

  List<XSSFCellFill> fills(StylesTable stylesTable) {
    Objects.requireNonNull(stylesTable, "stylesTable must not be null");
    return stylesTable.getFills();
  }

  int appendFill(StylesTable stylesTable, XSSFCellFill fill) {
    Objects.requireNonNull(stylesTable, "stylesTable must not be null");
    Objects.requireNonNull(fill, "fill must not be null");
    List<XSSFCellFill> fills = mutableFills(stylesTable);
    fills.add(fill);
    return fills.size() - 1;
  }

  @SuppressWarnings("unchecked")
  private List<XSSFCellFill> mutableFills(StylesTable stylesTable) {
    return (List<XSSFCellFill>) FILLS_FIELD.get(stylesTable);
  }

  static VarHandle requireFillsField(MethodHandles.Lookup lookup) {
    try {
      return MethodHandles.privateLookupIn(StylesTable.class, lookup)
          .findVarHandle(StylesTable.class, "fills", List.class);
    } catch (ReflectiveOperationException exception) {
      throw new IllegalStateException("Failed to access POI fill registry", exception);
    }
  }
}
