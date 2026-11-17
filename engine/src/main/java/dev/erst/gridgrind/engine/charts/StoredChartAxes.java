package dev.erst.gridgrind.engine.charts;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import javax.xml.namespace.QName;
import org.apache.xmlbeans.XmlObject;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCrosses;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDouble;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt;

/** Reads only referenced native axis facts without XDDF accessors that create declarations. */
final class StoredChartAxes {
  private static final String CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";
  private final StoredChartReader reader;

  // Construction confines writes to this read invocation; native family order determines ID
  // ownership.
  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private final Map<Long, Source> axesById = new LinkedHashMap<>();

  StoredChartAxes(CTPlotArea area, StoredChartReader reader) {
    this.reader = reader;
    // Preserve POI's existing family order and last-owner behavior for duplicate IDs.
    area.getCatAxList().forEach(axis -> index(axis, ExcelChartAxisKind.CATEGORY));
    area.getDateAxList().forEach(axis -> index(axis, ExcelChartAxisKind.DATE));
    area.getSerAxList().forEach(axis -> index(axis, ExcelChartAxisKind.SERIES));
    area.getValAxList().forEach(axis -> index(axis, ExcelChartAxisKind.VALUE));
  }

  List<ChartReport.Axis> selected(List<CTUnsignedInt> references) {
    List<ChartReport.Axis> selected = new ArrayList<>();
    for (CTUnsignedInt reference : references) {
      Source source = axesById.get(id(reference));
      if (source != null) selected.add(snapshot(source));
    }
    return List.copyOf(selected);
  }

  private void index(XmlObject axis, ExcelChartAxisKind kind) {
    axesById.put(id((CTUnsignedInt) required(axis, "axId")), new Source(axis, kind));
  }

  private long id(CTUnsignedInt id) {
    reader.requireValid(id.xgetVal(), "chart.axis.id");
    return id.getVal();
  }

  private ChartReport.Axis snapshot(Source source) {
    var position = (CTAxPos) required(source.xml(), "axPos");
    reader.requireValid(position.xgetVal(), "chart.axis.position");
    ExcelChartAxisPosition location =
        switch (position.getVal().toString()) {
          case "b" -> ExcelChartAxisPosition.BOTTOM;
          case "l" -> ExcelChartAxisPosition.LEFT;
          case "r" -> ExcelChartAxisPosition.RIGHT;
          case "t" -> ExcelChartAxisPosition.TOP;
          default -> throw new IllegalStateException("Unrecognized native axis position");
        };
    var deletion = (CTBoolean) child(source.xml(), "delete");
    if (deletion != null) reader.requireValid(deletion.xgetVal(), "chart.axis.delete");
    return new ChartReport.Axis(
        source.kind(), location, crossing(source.xml()), deletion == null || !deletion.getVal());
  }

  private ChartReport.AxisCrossing crossing(XmlObject source) {
    reader.requireSingleSelection(source, "chart.axis.crossing", Set.of("crosses", "crossesAt"));
    var mode = (CTCrosses) child(source, "crosses");
    if (mode != null) {
      reader.requireValid(mode.xgetVal(), "chart.axis.crosses");
      return new ChartReport.AxisCrossing.Mode(
          switch (mode.getVal().toString()) {
            case "autoZero" -> ExcelChartAxisCrosses.AUTO_ZERO;
            case "min" -> ExcelChartAxisCrosses.MIN;
            case "max" -> ExcelChartAxisCrosses.MAX;
            default -> throw new IllegalStateException("Unrecognized native axis crossing mode");
          });
    }
    var value = (CTDouble) child(source, "crossesAt");
    if (value == null) return new ChartReport.AxisCrossing.Missing();
    reader.requireValid(value.xgetVal(), "chart.axis.crossesAt");
    try (var cursor = value.newCursor()) {
      return new ChartReport.AxisCrossing.Value(
          Objects.requireNonNull(cursor.getAttributeText(new QName("val"))));
    }
  }

  private XmlObject required(XmlObject source, String name) {
    XmlObject child = child(source, name);
    if (child == null) throw reader.invalid("chart.axis." + name, null);
    return child;
  }

  private @Nullable XmlObject child(XmlObject source, String name) {
    reader.requireSingleSelection(source, "chart.axis." + name, Set.of(name));
    XmlObject[] children = source.selectChildren(new QName(CHART, name));
    return children.length == 0 ? null : children[0];
  }

  /** Native axis and its element-declared family, indexed before selecting referenced facts. */
  private record Source(XmlObject xml, ExcelChartAxisKind kind) {}
}
