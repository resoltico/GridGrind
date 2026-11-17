package dev.erst.gridgrind.engine.charts;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.core.problem.WorkbookContentLocation;
import dev.erst.gridgrind.engine.backend.xssf.UnexpectedNativeContentException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.Supplier;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.xmlbeans.XmlAnySimpleType;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlUnsignedInt;
import org.apache.xmlbeans.impl.values.XmlValueOutOfRangeException;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt;

/** Reads stored chart declarations without live resolution, padding, evaluation or cache writes. */
final class StoredChartReader {
  private static final String CHART_NAMESPACE =
      "http://schemas.openxmlformats.org/drawingml/2006/chart";
  private static final Set<String> AXIS_PLOTS =
      Set.of(
          "areaChart",
          "area3DChart",
          "barChart",
          "bar3DChart",
          "lineChart",
          "line3DChart",
          "radarChart",
          "scatterChart",
          "surfaceChart",
          "surface3DChart");
  private final Optional<String> partName;

  StoredChartReader(Optional<String> partName) {
    this.partName = Objects.requireNonNull(partName);
  }

  <T> T read(String field, Supplier<T> operation) {
    try {
      return operation.get();
    } catch (WorkbookContentRefusalException | UnexpectedNativeContentException exception) {
      throw exception;
    } catch (XmlValueOutOfRangeException exception) {
      throw invalid(field, exception);
    } catch (RuntimeException exception) {
      throw new UnexpectedNativeContentException(exception);
    }
  }

  String name(XSSFGraphicFrame frame) {
    return read(
        "chart.name",
        () -> {
          var properties = frame.getCTGraphicalObjectFrame().getNvGraphicFramePr();
          if (properties == null || properties.getCNvPr() == null)
            throw invalid("chart.name", null);
          var scalar = properties.getCNvPr().xgetName();
          requireValid(scalar, "chart.name");
          return scalar.getStringValue();
        });
  }

  void requireValid(@Nullable XmlAnySimpleType scalar, String field) {
    if (scalar == null || !scalar.validate()) throw invalid(field, null);
  }

  /** Refuse only native omissions that the pinned XDDF read path would repair by writing XML. */
  void requireReadPreconditions(CTPlotArea area) {
    try (var cursor = area.newCursor()) {
      for (boolean child = cursor.toFirstChild(); child; child = cursor.toNextSibling()) {
        var name = cursor.getName();
        if (name == null || !CHART_NAMESPACE.equals(name.getNamespaceURI())) continue;
        XmlObject plot = cursor.getObject();
        if (AXIS_PLOTS.contains(name.getLocalPart())
            && plot.selectChildren(CHART_NAMESPACE, "axId").length == 0) {
          throw invalid("chart.plot.axisReferences", null);
        }
        if ("barChart".equals(name.getLocalPart()) || "bar3DChart".equals(name.getLocalPart())) {
          requireReadField(plot, "barDir");
        } else if ("scatterChart".equals(name.getLocalPart())) {
          requireReadField(plot, "scatterStyle");
        }
      }
    }
  }

  private void requireReadField(XmlObject plot, String name) {
    String field = "chart.plot." + name;
    requireSingleSelection(plot, field, Set.of(name));
    XmlObject[] values = plot.selectChildren(CHART_NAMESPACE, name);
    if (values.length == 0 || !values[0].validate()) throw invalid(field, null);
  }

  void requireFields(XmlObject source, String... names) {
    for (String name : names) {
      String field = "chart." + name;
      requireSingleSelection(source, field, Set.of(name));
      for (XmlObject value : source.selectChildren(CHART_NAMESPACE, name)) {
        if (!value.validate()) throw invalid(field, null);
      }
    }
  }

  void requireMarker(XmlObject series) {
    requireSingleSelection(series, "chart.series.marker", Set.of("marker"));
    for (XmlObject marker : series.selectChildren(CHART_NAMESPACE, "marker")) {
      requireFields(marker, "size", "symbol");
    }
  }

  ChartReport.DataSource source(@Nullable CTAxDataSource source) {
    return read(
        "chart.series.categorySource",
        () -> {
          if (source == null) return new ChartReport.DataSource.Missing();
          requireSingleSelection(
              source,
              "chart.series.categorySource.choice",
              Set.of("numRef", "numLit", "strRef", "strLit", "multiLvlStrRef"));
          if (source.isSetNumRef()) return numericReference(source.getNumRef());
          if (source.isSetNumLit())
            return new ChartReport.DataSource.NumericLiteral(numeric(source.getNumLit()));
          if (source.isSetStrRef()) return stringReference(source.getStrRef());
          if (source.isSetStrLit())
            return new ChartReport.DataSource.StringLiteral(strings(source.getStrLit()));
          if (source.isSetMultiLvlStrRef())
            return ChartReport.DataSource.Unsupported.unmodeled("MULTI_LEVEL_STRING_REFERENCE");
          return new ChartReport.DataSource.Missing();
        });
  }

  ChartReport.DataSource source(@Nullable CTNumDataSource source) {
    return read(
        "chart.series.numericSource",
        () -> {
          if (source == null) return new ChartReport.DataSource.Missing();
          requireSingleSelection(
              source, "chart.series.numericSource.choice", Set.of("numRef", "numLit"));
          if (source.isSetNumRef()) return numericReference(source.getNumRef());
          if (source.isSetNumLit())
            return new ChartReport.DataSource.NumericLiteral(numeric(source.getNumLit()));
          return new ChartReport.DataSource.Missing();
        });
  }

  ChartReport.Title title(XSSFChart chart) {
    return read(
        "chart.title",
        () -> {
          requireSingleSelection(chart.getCTChart(), "chart.title", Set.of("title"));
          if (!chart.getCTChart().isSetTitle()) return new ChartReport.Title.None();
          requireSingleSelection(chart.getCTChart().getTitle(), "chart.title.text", Set.of("tx"));
          if (!chart.getCTChart().getTitle().isSetTx()) return new ChartReport.Title.None();
          var text = chart.getCTChart().getTitle().getTx();
          requireSingleSelection(text, "chart.title.choice", Set.of("strRef", "rich"));
          if (text.isSetStrRef()) return formulaTitle(text.getStrRef());
          if (!text.isSetRich()) return new ChartReport.Title.None();
          var stored = chart.getTitleText();
          return new ChartReport.Title.Text(stored == null ? "" : stored.getString());
        });
  }

  ChartReport.Title title(@Nullable CTSerTx title) {
    return read(
        "chart.series.title",
        () -> {
          if (title == null) return new ChartReport.Title.None();
          requireSingleSelection(title, "chart.series.title.choice", Set.of("strRef", "v"));
          if (title.isSetStrRef()) return formulaTitle(title.getStrRef());
          return title.isSetV()
              ? new ChartReport.Title.Text(title.getV())
              : new ChartReport.Title.None();
        });
  }

  /** Count native direct children before singleton getters can silently discard duplicates. */
  void requireSingleSelection(XmlObject source, String field, Set<String> alternatives) {
    int count = 0;
    try (var cursor = source.newCursor()) {
      if (!cursor.toFirstChild()) return;
      do {
        var name = cursor.getName();
        if (name != null
            && "http://schemas.openxmlformats.org/drawingml/2006/chart"
                .equals(name.getNamespaceURI())
            && alternatives.contains(name.getLocalPart())) {
          count++;
          if (count > 1) throw invalid(field, null);
        }
      } while (cursor.toNextSibling());
    }
  }

  private void requireStringReference(CTStrRef source) {
    requireSingleSelection(source, "chart.stringSource.formula", Set.of("f"));
    requireSingleSelection(source, "chart.stringSource.cache", Set.of("strCache"));
  }

  private ChartReport.Title formulaTitle(CTStrRef source) {
    requireStringReference(source);
    return new ChartReport.Title.Formula(
        Optional.ofNullable(source.getF()),
        source.isSetStrCache()
            ? strings(source.getStrCache())
            : new ChartReport.StringCache.Missing());
  }

  private ChartReport.DataSource stringReference(CTStrRef source) {
    requireStringReference(source);
    return new ChartReport.DataSource.StringReference(
        Optional.ofNullable(source.getF()),
        source.isSetStrCache()
            ? strings(source.getStrCache())
            : new ChartReport.StringCache.Missing());
  }

  private ChartReport.DataSource numericReference(CTNumRef source) {
    requireSingleSelection(source, "chart.numericSource.formula", Set.of("f"));
    requireSingleSelection(source, "chart.numericSource.cache", Set.of("numCache"));
    return new ChartReport.DataSource.NumericReference(
        Optional.ofNullable(source.getF()),
        source.isSetNumCache()
            ? numeric(source.getNumCache())
            : new ChartReport.NumericCache.Missing());
  }

  private ChartReport.StringCache.Stored strings(CTStrData data) {
    requireSingleSelection(data, "chart.stringCache.declaredCount", Set.of("ptCount"));
    List<ChartReport.StringPoint> points =
        data.getPtList().stream()
            .map(point -> {
              requireSingleSelection(point, "chart.stringCache.point.value", Set.of("v"));
              return new ChartReport.StringPoint(
                  unsigned(point.xgetIdx(), "chart.stringCache.point.index"),
                  Optional.ofNullable(point.getV()));
            })
            .toList();
    return new ChartReport.StringCache.Stored(
        count(data.getPtCount(), "chart.stringCache.declaredCount"), points);
  }

  private ChartReport.NumericCache.Stored numeric(CTNumData data) {
    requireSingleSelection(data, "chart.numericCache.declaredCount", Set.of("ptCount"));
    requireSingleSelection(data, "chart.numericCache.formatCode", Set.of("formatCode"));
    List<ChartReport.NumericPoint> points =
        data.getPtList().stream()
            .map(point -> {
              requireSingleSelection(point, "chart.numericCache.point.value", Set.of("v"));
              return new ChartReport.NumericPoint(
                  unsigned(point.xgetIdx(), "chart.numericCache.point.index"),
                  Optional.ofNullable(point.getV()),
                  point.isSetFormatCode() ? Optional.of(point.getFormatCode()) : Optional.empty());
            })
            .toList();
    return new ChartReport.NumericCache.Stored(
        count(data.getPtCount(), "chart.numericCache.declaredCount"),
        data.isSetFormatCode() ? Optional.of(data.getFormatCode()) : Optional.empty(),
        points);
  }

  private Optional<Long> count(@Nullable CTUnsignedInt count, String field) {
    return count == null ? Optional.empty() : unsigned(count.xgetVal(), field);
  }

  private Optional<Long> unsigned(@Nullable XmlUnsignedInt value, String field) {
    if (value == null) return Optional.empty();
    if (!value.validate()) throw invalid(field, null);
    return Optional.of(value.getLongValue());
  }

  WorkbookContentRefusalException invalid(String field, @Nullable Throwable cause) {
    return new WorkbookContentRefusalException(
        WorkbookContentIssue.INVALID_NATIVE_VALUE,
        Optional.of(
            new WorkbookContentLocation(
                Optional.empty(),
                partName,
                Optional.of(field),
                Optional.empty(),
                Optional.empty())),
        cause);
  }
}
