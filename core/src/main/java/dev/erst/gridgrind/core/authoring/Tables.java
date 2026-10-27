package dev.erst.gridgrind.core.authoring;

import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.Objects;

/** Focused table builders for the fluent Java authoring surface. */
public final class Tables {
  /** Table-style wrapper for fluent authoring. */
  public sealed interface Style permits NoStyle, NamedStyle {}

  /** Explicit no-style table marker. */
  public record NoStyle() implements Style {}

  /** Named table style with stripe and emphasis flags. */
  public record NamedStyle(
      String name,
      boolean showFirstColumn,
      boolean showLastColumn,
      boolean showRowStripes,
      boolean showColumnStripes)
      implements Style {
    public NamedStyle {
      Objects.requireNonNull(name, "name must not be null");
    }
  }

  /** Table-definition wrapper for fluent authoring. */
  public record Definition(
      TableName name, SheetName sheetName, RangeRef range, boolean totalsRowShown, Style style) {
    public Definition {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(range, "range must not be null");
      Objects.requireNonNull(style, "style must not be null");
    }
  }

  private Tables() {}

  /** Returns an explicit no-style table marker. */
  public static Style noStyle() {
    return new NoStyle();
  }

  /** Returns a named table style wrapper. */
  public static Style namedStyle(
      String name,
      boolean showFirstColumn,
      boolean showLastColumn,
      boolean showRowStripes,
      boolean showColumnStripes) {
    return new NamedStyle(name, showFirstColumn, showLastColumn, showRowStripes, showColumnStripes);
  }

  /** Returns a focused table definition. */
  public static Definition define(
      TableName name, SheetName sheetName, RangeRef range, boolean totalsRowShown, Style style) {
    return new Definition(name, sheetName, range, totalsRowShown, style);
  }

  static TableSpec toTableInput(Definition definition) {
    Definition nonNullDefinition =
        Objects.requireNonNull(definition, "definition must not be null");
    return TableSpec.withDefaultMetadata(
        nonNullDefinition.name(),
        nonNullDefinition.sheetName(),
        nonNullDefinition.range(),
        nonNullDefinition.totalsRowShown(),
        toTableStyleInput(nonNullDefinition.style()));
  }

  private static TableStyleSpec toTableStyleInput(Style style) {
    return switch (Objects.requireNonNull(style, "style must not be null")) {
      case NoStyle _ -> new TableStyleSpec.None();
      case NamedStyle named ->
          new TableStyleSpec.Named(
              named.name(),
              named.showFirstColumn(),
              named.showLastColumn(),
              named.showRowStripes(),
              named.showColumnStripes());
    };
  }
}
