package dev.erst.gridgrind.engine.charts;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.problem.WorkbookContentIssue;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookContentRefusalException;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

/** Direct native XML challenges singleton selection without imposing whole-chart validity. */
class StoredChartChoiceTest {
  private final StoredChartReader reader = new StoredChartReader(Optional.empty());

  @Test
  void repeatedSameChoiceChildrenRefuseRatherThanSelectingTheFirst() throws Exception {
    for (String choice : List.of("numRef", "numLit", "strRef", "strLit", "multiLvlStrRef")) {
      var source = CTAxDataSource.Factory.parse(fragment("<c:" + choice + "/><c:" + choice + "/>"));
      String before = source.xmlText();
      assertInvalid(() -> reader.source(source));
      assertEquals(before, source.xmlText());
    }
    for (String choice : List.of("numRef", "numLit")) {
      var source =
          CTNumDataSource.Factory.parse(fragment("<c:" + choice + "/><c:" + choice + "/>"));
      assertInvalid(() -> reader.source(source));
    }
    for (String choice : List.of("strRef", "v")) {
      var title = CTSerTx.Factory.parse(fragment("<c:" + choice + "/><c:" + choice + "/>"));
      assertInvalid(() -> reader.title(title));
    }
  }

  @Test
  void competingReportedFormulaCacheAndPointSingletonsRefuse() throws Exception {
    for (String body :
        List.of(
            "<c:numRef><c:f>First</c:f><c:f>Second</c:f></c:numRef>",
            "<c:numRef><c:numCache/><c:numCache/></c:numRef>",
            "<c:strRef><c:f>First</c:f><c:f>Second</c:f></c:strRef>",
            "<c:strRef><c:strCache/><c:strCache/></c:strRef>",
            "<c:numLit><c:ptCount val=\"1\"/><c:ptCount val=\"2\"/></c:numLit>",
            "<c:strLit><c:ptCount val=\"1\"/><c:ptCount val=\"2\"/></c:strLit>",
            "<c:numLit><c:formatCode>0</c:formatCode><c:formatCode>General</c:formatCode></c:numLit>",
            "<c:numLit><c:pt idx=\"0\"><c:v>1</c:v><c:v>2</c:v></c:pt></c:numLit>",
            "<c:strLit><c:pt idx=\"0\"><c:v>first</c:v><c:v>second</c:v></c:pt></c:strLit>")) {
      var source = CTAxDataSource.Factory.parse(fragment(body));
      String before = source.xmlText();
      assertInvalid(() -> reader.source(source));
      assertEquals(before, source.xmlText());
    }
  }

  @Test
  void absentSingleUnmodeledAndRepeatedPointIndicesRemainFactual() throws Exception {
    assertInstanceOf(
        ChartReport.DataSource.Missing.class, reader.source(CTAxDataSource.Factory.newInstance()));
    assertInstanceOf(
        ChartReport.DataSource.NumericReference.class,
        reader.source(CTAxDataSource.Factory.parse(fragment("<c:numRef/>"))));
    assertInstanceOf(
        ChartReport.DataSource.Unsupported.class,
        reader.source(CTAxDataSource.Factory.parse(fragment("<c:multiLvlStrRef/>"))));
    var source =
        CTAxDataSource.Factory.parse(
            fragment(
                "<c:numLit><c:pt idx=\"2\"><c:v>#N/A</c:v></c:pt><c:pt"
                    + " idx=\"2\"><c:v/></c:pt><c:pt/></c:numLit>"));
    String before = source.xmlText();
    var literal =
        assertInstanceOf(ChartReport.DataSource.NumericLiteral.class, reader.source(source));
    var stored = assertInstanceOf(ChartReport.NumericCache.Stored.class, literal.data());
    assertEquals(
        List.of(
            new ChartReport.NumericPoint(Optional.of(2L), Optional.of("#N/A"), Optional.empty()),
            new ChartReport.NumericPoint(Optional.of(2L), Optional.of(""), Optional.empty()),
            new ChartReport.NumericPoint(Optional.empty(), Optional.empty(), Optional.empty())),
        stored.points());
    assertEquals(before, source.xmlText());
  }

  @ParameterizedTest
  @ValueSource(ints = {0, 1, 2})
  void chartTitleAndTextSingletonsRefuseWithoutMutatingNativeXml(int scenario)
      throws java.io.IOException {
    try (var workbook = new XSSFWorkbook()) {
      var chart =
          workbook
              .createSheet("Title " + scenario)
              .createDrawingPatriarch()
              .createChart(new XSSFClientAnchor());
      var title = chart.getCTChart().addNewTitle();
      var text = title.addNewTx();
      text.addNewStrRef().setF("First");
      if (scenario == 0) chart.getCTChart().addNewTitle();
      if (scenario == 1) title.addNewTx();
      if (scenario == 2) text.addNewStrRef().setF("Second");
      String before = chart.getCTChartSpace().xmlText();
      assertInvalid(() -> reader.title(chart));
      assertEquals(before, chart.getCTChartSpace().xmlText());
    }
  }

  private static void assertInvalid(org.junit.jupiter.api.function.Executable operation) {
    assertEquals(
        WorkbookContentIssue.INVALID_NATIVE_VALUE,
        assertThrows(WorkbookContentRefusalException.class, operation).issue());
  }

  private static String fragment(String body) {
    return "<xml-fragment xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
        + body
        + "</xml-fragment>";
  }
}
