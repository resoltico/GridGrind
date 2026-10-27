package dev.erst.gridgrind.engine.poi;

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
@SuppressWarnings("PMD.CommentRequired")
public class FillRegistryAccess {
  public static final PrivateContract FILLS_FIELD_CONTRACT =
      PrivateContract.field(StylesTable.class, "fills", "gradient-fill registry synchronization");
  private static final VarHandle FILLS_FIELD = requireFillsField(MethodHandles.lookup());

  public static FillRegistryAccess poiApi() {
    return new FillRegistryAccess();
  }

  public List<XSSFCellFill> fills(StylesTable stylesTable) {
    Objects.requireNonNull(stylesTable, "stylesTable must not be null");
    return stylesTable.getFills();
  }

  public int appendFill(StylesTable stylesTable, XSSFCellFill fill) {
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

  public static VarHandle requireFillsField(MethodHandles.Lookup lookup) {
    return PoiAccess.requireVarHandle(lookup, FILLS_FIELD_CONTRACT, List.class);
  }
}
